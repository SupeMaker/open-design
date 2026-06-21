// Plugin snapshot resolver — wires the pure `applyPlugin()` into the
// daemon's `POST /api/projects` and `POST /api/runs` paths. Spec §8.2.1
// invariant I3: the AppliedPluginSnapshot is the only contract between
// "plugin" and "run". This module owns the side-effect-bearing edges:
//
//   1. Caller supplies `appliedPluginSnapshotId` → look it up, verify it
//      isn't stale.
//   2. Caller supplies `pluginId` (+ optional `pluginInputs`,
//      `grantCaps`) → run `applyPlugin()` with the live registry,
//      persist a `createSnapshot()` row, and return the new snapshot id.
//   3. Neither field present → return `null`; the caller proceeds with
//      the legacy non-plugin code path.
//
// Capability gating: when the resolved snapshot is `restricted` and any
// `capabilitiesRequired` is missing from `capabilitiesGranted`, we
// short-circuit with the §9.1 / exit-66 / 409 body. The caller maps the
// returned `error` shape to either an HTTP 409 or a stderr JSON envelope.
//
// This module is the single entry point for both project create and
// run start; all snapshot wiring goes through here so the behavior stays
// deterministic across CLI / desktop / web.

import type Database from 'better-sqlite3';
import type {
  AppliedPluginSnapshot,
  ApplyResult,
  InstalledPluginRecord,
  PluginConnectorBinding,
} from '@open-design/contracts';
import {
  applyPlugin,
  MissingInputError,
  type ApplyTrust,
} from './apply.js';
import {
  getInstalledPlugin,
} from './registry.js';
import {
  createSnapshot,
  getSnapshot,
  linkSnapshotToConversation,
  linkSnapshotToProject,
  linkSnapshotToRun,
} from './snapshots.js';
import { getManifestContextCraft } from './context-craft.js';
import {
  type ConnectorProbe,
} from './connector-gate.js';
import type { RegistryView } from '@open-design/plugin-runtime';

type SqliteDb = Database.Database;

export interface ResolveSnapshotInput {
  db: SqliteDb;
  body: Record<string, unknown> | null | undefined;
  // The project this snapshot will pin to. For the run-create path we
  // always know it (the run carries `projectId`). For project-create we
  // pass the freshly-inserted project id.
  projectId: string;
  conversationId?: string | null | undefined;
  runId?: string | null | undefined;
  // Pluggable for tests; in production these are the daemon's live
  // skill / design-system catalogs (server.ts wires them).
  registry: RegistryView;
  connectorProbe?: ConnectorProbe | undefined;
  // Optional active-project DS binding. Forwarded to `applyPlugin` so
  // plugins that declared `od.context.designSystem.primary: true` get
  // bound to the project's DS at apply time.
  activeProjectDesignSystem?: { id: string; title?: string } | undefined;
}

export interface ResolveSnapshotOk {
  ok: true;
  snapshotId: string;
  snapshot: AppliedPluginSnapshot;
  applyResult?: ApplyResult;
  // Whether this call created a new snapshot (true) or reused an
  // explicit `appliedPluginSnapshotId` (false). Used by callers to
  // decide when to re-link to a different project / run / conversation.
  created: boolean;
}

export interface ResolveSnapshotError {
  ok: false;
  status: number; // HTTP status to return
  exitCode: number; // Matching CLI exit code (§12.4)
  body: {
    error: {
      code: string;
      message: string;
      data?: Record<string, unknown>;
    };
  };
}

export type ResolveSnapshotResult = ResolveSnapshotOk | ResolveSnapshotError | null;

// Read the snapshot id that's currently pinned on a project row (if any).
// Returns null when the project is missing or has no snapshot pinned.
// Used by resolvePluginSnapshot's fallback so a plain `POST /api/runs
// { projectId }` reuses the snapshot the user picked at project create
// time — without forcing every caller to re-thread the snapshot id.
function readProjectPinnedSnapshotId(db: SqliteDb, projectId: string): string | null {
  try {
    const row = db
      .prepare(`SELECT applied_plugin_snapshot_id AS id FROM projects WHERE id = ?`)
      .get(projectId) as { id?: string | null } | undefined;
    const id = row?.id;
    return typeof id === 'string' && id.length > 0 ? id : null;
  } catch {
    return null;
  }
}

// Pull plugin-bearing fields off the request body without mutating it.
function pickPluginFields(body: Record<string, unknown> | null | undefined) {
  if (!body || typeof body !== 'object') return {};
  const pluginId = typeof body.pluginId === 'string' && body.pluginId.trim().length > 0
    ? body.pluginId.trim()
    : undefined;
  const snapshotId = typeof body.appliedPluginSnapshotId === 'string'
    && body.appliedPluginSnapshotId.trim().length > 0
    ? body.appliedPluginSnapshotId.trim()
    : undefined;
  const pluginInputs =
    body.pluginInputs && typeof body.pluginInputs === 'object'
      ? (body.pluginInputs as Record<string, unknown>)
      : body.inputs && typeof body.inputs === 'object'
        ? (body.inputs as Record<string, unknown>)
        : {};
  const grantCaps = Array.isArray(body.grantCaps)
    ? (body.grantCaps as unknown[])
        .filter((c): c is string => typeof c === 'string')
    : [];
  const locale = typeof body.locale === 'string' ? body.locale : undefined;
  return { pluginId, snapshotId, pluginInputs, grantCaps, locale };
}

export function resolvePluginSnapshot(input: ResolveSnapshotInput): ResolveSnapshotResult {
  // 从请求体中提取插件相关字段
  const fields = pickPluginFields(input.body);
  
  // 如果调用方在请求体中没有指定插件/快照，但项目上已经固定了一个快照
  // （由之前创建项目/对话时运行插件设置的），则复用它。
  // 这就是为什么 ChatComposer 的"开始运行"路径在用户在 NewProjectPanel 中
  // 选择了插件后能正常工作的原因 —— 请求体只携带了 `projectId`。
  if (!fields.pluginId && !fields.snapshotId && input.projectId) {
    // 从数据库中读取项目固定（绑定）的快照 ID
    const pinned = readProjectPinnedSnapshotId(input.db, input.projectId);
    if (pinned) {
      // 如果有固定的快照，将其设置为要使用的快照 ID
      fields.snapshotId = pinned;
    }
  }
  
  // 如果既没有插件 ID 也没有快照 ID，直接返回 null（无需解析）
  if (!fields.pluginId && !fields.snapshotId) return null;

  // 路径1：显式指定了快照 ID —— 查找快照并验证其状态。
  if (fields.snapshotId) {
    // 从数据库获取快照对象
    const snapshot = getSnapshot(input.db, fields.snapshotId);
    if (!snapshot) {
      // 快照未找到，返回 404 错误
      return {
        ok: false, // 操作失败
        status: 404, // HTTP 状态码：未找到
        exitCode: 65, // 退出码：数据不可用
        body: {
          error: {
            code: 'snapshot-not-found', // 错误代码：快照未找到
            message: `Applied plugin snapshot ${fields.snapshotId} not found`, // 错误消息
            data: { snapshotId: fields.snapshotId }, // 附加数据：快照 ID
          },
        },
      };
    }
    if (snapshot.status === 'stale') {
      // 快照已标记为过期，返回 409 错误
      return {
        ok: false, // 操作失败
        status: 409, // HTTP 状态码：冲突
        exitCode: 72, // 退出码：资源已过期
        body: {
          error: {
            code: 'snapshot-stale', // 错误代码：快照已过期
            message: `Snapshot ${fields.snapshotId} was marked stale; re-apply the plugin or replay the run.`, // 错误消息：提示重新应用插件或重放运行
            data: {
              snapshotId: snapshot.snapshotId, // 快照 ID
              pluginId: snapshot.pluginId, // 插件 ID
              snapshotVersion: snapshot.pluginVersion, // 快照版本
            },
          },
        },
      };
    }
    // 快照有效，返回成功结果，created 为 false（非新创建）
    return finalizeOk({
      input,
      snapshot,
      created: false, // 标记这不是新创建的快照
    });
  }

  // 路径2：指定了插件 ID —— 运行插件的 apply 逻辑，持久化生成一个新快照。
  // 从数据库中获取已安装的插件
  const plugin = getInstalledPlugin(input.db, fields.pluginId!);
  if (!plugin) {
    // 插件未安装，返回 404 错误
    return {
      ok: false, // 操作失败
      status: 404, // HTTP 状态码：未找到
      exitCode: 65, // 退出码：数据不可用
      body: {
        error: {
          code: 'plugin-not-found', // 错误代码：插件未找到
          message: `Plugin "${fields.pluginId}" is not installed.`, // 错误消息：插件未安装
          data: { pluginId: fields.pluginId }, // 附加数据：插件 ID
        },
      },
    };
  }

  let applyComputed; // 存储应用插件的计算结果
  try {
    // 执行插件的 apply 函数，传入插件实例和相关参数
    applyComputed = applyPlugin({
      plugin, // 插件对象
      inputs: fields.pluginInputs ?? {}, // 插件输入参数，默认为空对象
      registry: input.registry, // 注册表
      activeProjectDesignSystem: input.activeProjectDesignSystem, // 当前项目激活的设计系统
      connectorProbe: input.connectorProbe, // 连接器探针
      locale: fields.locale, // 语言区域设置
    });
  } catch (err) {
    // 捕获插件应用过程中的错误
    if (err instanceof MissingInputError) {
      // 如果是缺少必需输入的错误，返回 422 错误
      return {
        ok: false, // 操作失败
        status: 422, // HTTP 状态码：无法处理的实体
        exitCode: 67, // 退出码：用户输入错误
        body: {
          error: {
            code: 'missing-input', // 错误代码：缺少输入
            message: `Plugin "${fields.pluginId}" is missing required inputs: ${err.fields.join(', ')}.`, // 错误消息：列出缺少的字段
            data: { pluginId: fields.pluginId, missing: err.fields }, // 附加数据：插件 ID 和缺少的字段列表
          },
        },
      };
    }
    // 其他未知错误，直接抛出
    throw err;
  }

  const result = applyComputed.result; // 获取应用插件的结果
  const trust: ApplyTrust = result.trust; // 获取信任级别
  // 合并插件授予的能力和调用方额外授权的能力，去重
  const grantedSet = new Set([...result.capabilitiesGranted, ...fields.grantCaps]);
  const merged = Array.from(grantedSet); // 转换为数组，得到最终授予的能力列表

  // 检查是否有必需但未被授予的能力
  const missing = result.capabilitiesRequired.filter((c) => !grantedSet.has(c));
  if (trust === 'restricted' && missing.length > 0) {
    // 如果是受限信任级别且存在未满足的能力要求，返回能力不足的错误
    return capabilitiesRequiredError({
      pluginId: plugin.id, // 插件 ID
      pluginVersion: plugin.version, // 插件版本
      required: result.capabilitiesRequired, // 插件要求的所有能力
      granted: merged, // 实际授予的能力
      missing, // 缺失的能力
    });
  }

  // 将应用结果持久化为一个新的快照记录
  const persisted = createSnapshot(input.db, {
    projectId: input.projectId, // 关联的项目 ID
    conversationId: input.conversationId ?? null, // 关联的对话 ID，可选
    runId: input.runId ?? null, // 关联的运行 ID，可选
    pluginId: result.appliedPlugin.pluginId, // 插件 ID
    pluginSpecVersion: result.appliedPlugin.pluginSpecVersion ?? plugin.manifest.specVersion, // 插件规范版本，优先使用结果中的，否则用清单中的
    pluginVersion: result.appliedPlugin.pluginVersion, // 插件版本
    pluginTitle: result.appliedPlugin.pluginTitle, // 插件标题
    pluginDescription: result.appliedPlugin.pluginDescription, // 插件描述
    manifestSourceDigest: applyComputed.manifestSourceDigest, // 清单源码摘要
    sourceMarketplaceId: result.appliedPlugin.sourceMarketplaceId ?? null, // 来源市场 ID
    sourceMarketplaceEntryName: result.appliedPlugin.sourceMarketplaceEntryName ?? null, // 来源市场条目名称
    sourceMarketplaceEntryVersion: result.appliedPlugin.sourceMarketplaceEntryVersion ?? null, // 来源市场条目版本
    marketplaceTrust: result.appliedPlugin.marketplaceTrust ?? null, // 市场信任级别
    resolvedSource: result.appliedPlugin.resolvedSource ?? null, // 解析后的来源
    resolvedRef: result.appliedPlugin.resolvedRef ?? null, // 解析后的引用
    archiveIntegrity: result.appliedPlugin.archiveIntegrity ?? null, // 归档完整性校验
    pinnedRef: result.appliedPlugin.pinnedRef ?? null, // 固定的引用
    taskKind: result.appliedPlugin.taskKind, // 任务类型
    inputs: result.appliedPlugin.inputs, // 输入参数
    resolvedContext: result.appliedPlugin.resolvedContext, // 解析后的上下文
    craftRequires: result.appliedPlugin.craftRequires ?? getManifestContextCraft(plugin.manifest), // Craft 要求，优先使用结果中的，否则从清单中提取
    pipeline: result.appliedPlugin.pipeline, // 流水线配置
    genuiSurfaces: result.appliedPlugin.genuiSurfaces ?? [], // 生成 UI 的表面配置，默认为空数组
    capabilitiesGranted: merged, // 最终授予的能力列表
    capabilitiesRequired: result.capabilitiesRequired, // 要求的能力列表
    assetsStaged: result.appliedPlugin.assetsStaged, // 暂存的资源
    connectorsRequired: result.appliedPlugin.connectorsRequired, // 需要的连接器
    connectorsResolved: result.appliedPlugin.connectorsResolved, // 已解析的连接器
    mcpServers: result.appliedPlugin.mcpServers, // MCP 服务器配置
    query: result.query, // 查询信息
  });

  // 返回成功结果，created 为 true（表示这是新创建的快照）
  return finalizeOk({
    input,
    snapshot: persisted, // 使用持久化后的快照
    applyResult: { ...result, appliedPlugin: persisted }, // 应用结果，将 appliedPlugin 替换为持久化后的快照
    created: true, // 标记这是新创建的快照
  });
}
function finalizeOk(args: {
  input: ResolveSnapshotInput;
  snapshot: AppliedPluginSnapshot;
  applyResult?: ApplyResult;
  created: boolean;
}): ResolveSnapshotOk {
  // Pin the snapshot to whichever surfaces the caller already knows.
  // Order matters: link to project (always) before conversation/run so
  // the foreign key is satisfied and `expires_at` clears in one statement.
  const { db } = args.input;
  const snap = args.snapshot;
  if (args.input.projectId) {
    linkSnapshotToProject(db, snap.snapshotId, args.input.projectId);
  }
  if (args.input.conversationId) {
    linkSnapshotToConversation(db, snap.snapshotId, args.input.conversationId);
  }
  if (args.input.runId) {
    linkSnapshotToRun(db, snap.snapshotId, args.input.runId);
  }
  return {
    ok: true,
    snapshotId: snap.snapshotId,
    snapshot: snap,
    ...(args.applyResult ? { applyResult: args.applyResult } : {}),
    created: args.created,
  };
}

export function capabilitiesRequiredError(args: {
  pluginId: string;
  pluginVersion: string;
  required: string[];
  granted: string[];
  missing: string[];
}): ResolveSnapshotError {
  const remediation = [
    `od plugin trust ${args.pluginId} --capabilities ${args.missing.join(',')}`,
    `or pass --grant-caps ${args.missing.join(',')} to the apply / run command`,
  ];
  return {
    ok: false,
    status: 409,
    exitCode: 66,
    body: {
      error: {
        code: 'capabilities-required',
        message: `Plugin ${args.pluginId} requires capabilities not yet granted.`,
        data: {
          pluginId: args.pluginId,
          pluginVersion: args.pluginVersion,
          required: args.required,
          granted: args.granted,
          missing: args.missing,
          remediation,
        },
      },
    },
  };
}

// Convenience pass-through so tests that already imported the helper
// don't need to reach into other files.
export type { PluginConnectorBinding };
