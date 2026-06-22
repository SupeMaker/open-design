import {
  useCallback,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
  useLayoutEffect,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react';
import { AnimatePresence } from 'motion/react';
import { createHtmlArtifactManifest, inferLegacyManifest } from '../artifacts/manifest';
import { resolveHtmlPointerArtifactTarget } from '../artifacts/pointer';
import { validateHtmlArtifact } from '../artifacts/validate';
import { recoverHtmlDocumentFromMarkdownFence, recoverStandaloneHtmlDocument, resolvePersistedArtifactHtml } from '../artifacts/recover';
import { createArtifactParser } from '../artifacts/parser';
import {
  findFirstQuestionForm,
  hasUnterminatedQuestionForm,
  parsePartialQuestionForm,
  type QuestionForm,
} from '../artifacts/question-form';
import { parseSubmittedAnswers } from './QuestionForm';
import { useI18n } from '../i18n';
import { streamMessage } from '../providers/anthropic';
import {
  fetchChatRunStatus,
  fetchVelaLoginStatus,
  listActiveChatRuns,
  listProjectRuns,
  reattachDaemonRun,
  reportChatRunFeedback,
  streamViaDaemon,
} from '../providers/daemon';
import { fetchElevenLabsVoiceOptions } from '../providers/elevenlabs-voices';
import { normalizeCustomReason } from '@open-design/contracts/analytics';
import {
  deletePreviewComment,
  fetchConnectorStatuses,
  fetchPreviewComments,
  fetchDesignSystem,
  fetchDesignTemplate,
  fetchProjectDesignSystemPackageAudit,
  fetchLiveArtifacts,
  fetchProjectFiles,
  fetchSkill,
  patchPreviewCommentStatus,
  projectRawUrl,
  uploadProjectFiles,
  upsertPreviewComment,
  writeProjectTextFile,
} from '../providers/registry';
import { useProjectFileEvents, type ProjectEvent } from '../providers/project-events';
import { claimRunTurnIndex } from '../analytics/identity';
import { useCoalescedCallback } from '../hooks/useCoalescedCallback';
import {
  composeSystemPrompt,
  type AudioVoiceOption,
  type MemorySystemPromptResponse,
  type ResearchOptions,
} from '@open-design/contracts';
import {
  anonymizeArtifactId,
  artifactKindToTracking,
  projectKindToTracking,
} from '@open-design/contracts/analytics';
import type {
  TrackingArtifactKind,
  TrackingDesignSystemApplyTargetKind,
  TrackingDesignSystemOrigin,
  TrackingDesignSystemStatusValue,
} from '@open-design/contracts/analytics';
import { useAnalytics } from '../analytics/provider';
import {
  trackArtifactHeaderClick,
  trackComposerBarClick,
  trackDesignSystemApplyResult,
  trackPageView,
  trackRunCreated,
  trackRunFinished,
} from '../analytics/events';
import {
  buildByokRunCreatedProps,
  buildByokRunFinishedProps,
} from '../analytics/byok-run';
import {
  clearOnboardingSessionId,
  peekOnboardingSessionId,
} from '../analytics/onboarding-session';
import { navigate } from '../router';
import { agentDisplayName, agentModelDisplayName } from '../utils/agentLabels';
import { isMacPlatform } from '../utils/platform';
import {
  canAutoRenameProjectFromPrompt,
  summarizeProjectNameFromPrompt,
} from '../utils/projectName';
import {
  apiProtocolAgentId,
  apiProtocolModelLabel,
  usesAnthropicProxy,
} from '../utils/apiProtocol';
import { playSound, showCompletionNotification } from '../utils/notifications';
import { randomUUID } from '../utils/uuid';
import { DEFAULT_NOTIFICATIONS } from '../state/config';
import type { TodoItem } from '../runtime/todos';
import { appendErrorStatusEvent } from '../runtime/chat-events';
import { RESUME_CONTINUE_PROMPT } from '../runtime/resume';
import {
  buildDesignSystemPackageAuditRepairPrompt,
  summarizeDesignSystemPackageAudit,
} from '../runtime/design-system-package-audit';
import { isLiveArtifactTabId, liveArtifactTabId } from '../types';
import {
  DESIGN_SYSTEM_WORKSPACE_DISPLAY_TITLE,
  isDesignSystemWorkspacePrompt,
} from '../design-system-auto-prompt';
import {
  createConversation,
  deleteConversation as deleteConversationApi,
  fetchAppliedPluginSnapshot,
  getTemplate,
  installGeneratedPluginFolder,
  listConversations,
  listMessages,
  loadTabs,
  patchConversation,
  patchProject,
  saveMessage,
  startGeneratedPluginShareTask,
  cacheTabsLocally,
  persistTabsToDaemonNow,
  listPlugins,
  type SaveMessageOptions,
  waitGeneratedPluginShareTask,
} from '../state/projects';
import type { AppliedPluginSnapshot, ChatAnalyticsEntryFrom, ChatSessionMode, InstalledPluginRecord, WorkspaceContextItem } from '@open-design/contracts';
import type {
  AgentEvent,
  AgentInfo,
  AppConfig,
  Artifact,
  ChatAttachment,
  ChatCommentAttachment,
  ChatMessage,
  ChatMessageFeedbackChange,
  Conversation,
  DesignSystemSummary,
  OpenTabsState,
  Project,
  ProjectMetadata,
  PreviewComment,
  PreviewCommentAttachment,
  PreviewCommentTarget,
  ProjectFile,
  ProjectTemplate,
  LiveArtifactEventItem,
  LiveArtifactSummary,
  SkillSummary,
} from '../types';
import { historyWithApiAttachmentContext } from '../api-attachment-context';
import {
  commentsToAttachments,
  historyWithCommentAttachmentContext,
  mergeAttachedComments,
  mergePreviewCommentAttachments,
  queuedSlideNavTarget,
  removeAttachedComment,
} from '../comments';
import { filterImplicitProducedFiles } from '../produced-files';
import { buildPptxExportPrompt } from '../lib/build-pptx-export-prompt';
import { AvatarMenu } from './AvatarMenu';
import { EntrySettingsMenu } from './EntrySettingsMenu';
import { HandoffButton } from './HandoffButton';
import { Icon } from './Icon';
import { DesignSystemPicker } from './DesignSystemPicker';
import { PluginDetailsModal } from './PluginDetailsModal';
import { DesignSystemPreviewModal } from './DesignSystemPreviewModal';
import { ChatPane } from './ChatPane';
import type { QuestionFormOpenRequest } from './AssistantMessage';
import type { ChatSendMeta } from './ChatComposer';
import {
  CritiqueTheaterMount,
  useCritiqueTheaterEnabled,
} from './Theater';
import { useIframeKeepAlivePool } from './IframeKeepAlivePool';
import {
  decideAutoOpenAfterWrite,
  selectAutoOpenProducedHtml,
} from './auto-open-file';
import { buildRepoImportPrompt, designSystemNeedsRepoConnect } from './design-system-github-evidence';
import { collectReferencedJsxNames } from '../runtime/jsx-module-refs';
import { FileWorkspace } from './FileWorkspace';
import {
  type PluginFolderAgentAction,
} from './design-files/pluginFolderActions';
import { SHARE_TO_COMMUNITY_PROMPT } from './share-to-community/shareToCommunityPrompt';
import { CenteredLoader } from './Loading';
import type { SettingsSection } from './SettingsDialog';
import { Toast } from './Toast';
import { useDesignMdState } from '../hooks/useDesignMdState';
import { useFinalizeProject } from '../hooks/useFinalizeProject';
import { useProjectDetail } from '../hooks/useProjectDetail';
import { useTerminalLaunch } from '../hooks/useTerminalLaunch';
import { buildContinueInCliToast } from '../lib/build-continue-in-cli-toast';
import { buildClipboardPrompt } from '../lib/build-clipboard-prompt';
import { copyToClipboard } from '../lib/copy-to-clipboard';
import { effectiveMaxTokens } from '../state/maxTokens';
import { effectiveAgentModelChoice } from './agentModelSelection';
import { mediaExecutionPolicyForProjectMetadata } from '../media/execution-policy';
import { mediaModelProviderId } from '../media/models';
import {
  useByokImageModelOptions,
  useByokVideoModelOptions,
  useByokSpeechModelOptions,
} from '../media/aihubmix-image-models';
import {
  buildFinalizeCredentialsMissingToast,
  buildFinalizeRequest,
} from '../lib/resolve-finalize-request';


type ProjectChatSendMeta = ChatSendMeta & {
  queueOnly?: boolean;
  retryOfAssistantId?: string;
  sessionMode?: ChatSessionMode;
  /** Overrides the run_created / run_finished `entry_from` analytics prop for
   *  this send (e.g. 'resume_continue' from the resumable-failure Continue
   *  action). Behavior never depends on it; it only shapes PostHog props. */
  entryFrom?: ChatAnalyticsEntryFrom;
};

export function mergeSavedPreviewComment(current: PreviewComment[], saved: PreviewComment): PreviewComment[] {
  const existingIndex = current.findIndex((comment) => comment.id === saved.id);
  if (existingIndex < 0) return [...current, saved];
  return current.map((comment, index) => (index === existingIndex ? saved : comment));
}

function mergeServerMessageWithLocal(server: ChatMessage, local?: ChatMessage): ChatMessage {
  if (!local) return server;
  const merged: ChatMessage = { ...server };
  if (local.role === 'assistant' && server.role === 'assistant') {
    if ((local.content?.length ?? 0) > (server.content?.length ?? 0)) {
      merged.content = local.content;
    }
    if ((local.events?.length ?? 0) > (server.events?.length ?? 0)) {
      merged.events = local.events;
    }
  }
  if (!server.producedFiles?.length && local.producedFiles?.length) {
    merged.producedFiles = local.producedFiles;
  }
  if (!server.preTurnFileNames?.length && local.preTurnFileNames?.length) {
    merged.preTurnFileNames = local.preTurnFileNames;
  }
  if (!server.lastRunEventId && local.lastRunEventId) {
    merged.lastRunEventId = local.lastRunEventId;
  }
  if (!server.startedAt && local.startedAt) {
    merged.startedAt = local.startedAt;
  }
  if (!server.endedAt && local.endedAt) {
    merged.endedAt = local.endedAt;
  }
  if (!server.runStatus && local.runStatus) {
    merged.runStatus = local.runStatus;
  }
  return merged;
}

export function mergeServerMessagesIntoConversation(
  current: ChatMessage[],
  serverMessages: ChatMessage[],
): ChatMessage[] {
  const currentById = new Map(current.map((message) => [message.id, message]));
  const serverIds = new Set(serverMessages.map((message) => message.id));
  const merged = serverMessages.map((message) =>
    mergeServerMessageWithLocal(message, currentById.get(message.id)),
  );
  for (const message of current) {
    if (!serverIds.has(message.id)) merged.push(message);
  }
  return merged;
}

interface Props {
  project: Project;
  routeFileName: string | null;
  /**
   * Routed conversation id. When set (the URL is
   * `/projects/:id/conversations/:cid[/...]`), the project view picks
   * this conversation as active instead of defaulting to `list[0]`.
   * Falls through to the default picker if the conversation does not
   * exist (e.g. the run was deleted between the route landing and the
   * conversation list loading). Issue #1505. Optional so existing
   * test harnesses that mount ProjectView with a stub props bag do
   * not have to be updated; production callers in `App.tsx` always
   * pass the value from `useRoute()`.
   */
  routeConversationId?: string | null;
  config: AppConfig;
  agents: AgentInfo[];
  // Mentionable functional skills — already filtered by config.disabledSkills
  // upstream, so this drives only the chat composer's @-picker scope. For
  // resolving an existing project's `skillId` (which can also point at a
  // design template after the skills/design-templates split) use
  // `designTemplates` as a fallback in composedSystemPrompt() and in the
  // skill-name / skill-mode lookups below.
  skills: SkillSummary[];
  // All known design templates (unfiltered). Required so projects created
  // from the Templates surface keep composing the template body in API
  // mode even when the user later disables the template in Settings.
  designTemplates: SkillSummary[];
  designSystems: DesignSystemSummary[];
  daemonLive: boolean;
  onModeChange: (mode: AppConfig['mode']) => void;
  onAgentChange: (id: string) => void;
  onAgentModelChange: (
    id: string,
    choice: { model?: string; reasoning?: string },
  ) => void;
  onApiModelChange?: (model: string) => void;
  onRefreshAgents: () => void;
  onThemeChange?: (theme: AppConfig['theme']) => void;
  onOpenSettings: (section?: SettingsSection) => void;
  onOpenAmrSettings?: () => void;
  onOpenMcpSettings?: () => void;
  onBrowsePlugins?: () => void;
  onOpenConnectors?: () => void;
  // Pet wiring forwarded to the chat composer so users can adopt /
  // wake / tuck a pet without leaving the project view.
  onAdoptPetInline?: (petId: string) => void;
  onTogglePet?: () => void;
  onOpenPetSettings?: () => void;
  onBack: () => void;
  onClearPendingPrompt: () => void;
  onTouchProject: () => void;
  onProjectChange: (next: Project) => void;
  onProjectsRefresh: () => void;
  onChangeDefaultDesignSystem?: (designSystemId: string | null) => void;
  onDesignSystemsRefresh?: () => Promise<void> | void;
}

interface QueuedChatSend {
  id: string;
  conversationId: string;
  prompt: string;
  attachments: ChatAttachment[];
  commentAttachments: ChatCommentAttachment[];
  meta?: ProjectChatSendMeta;
  createdAt: number;
}

interface QueuedChatSendUpdate {
  prompt: string;
  attachments: ChatAttachment[];
  commentAttachments: ChatCommentAttachment[];
  meta?: ChatSendMeta;
}

let liveArtifactEventSequence = 0;
const CHAT_PANEL_WIDTH_STORAGE_KEY = 'open-design.project.chatPanelWidth';
const DEFAULT_CHAT_PANEL_WIDTH = 460;
const MIN_CHAT_PANEL_WIDTH = 345;
const MAX_CHAT_PANEL_WIDTH = 720;
const COMMENT_INSPECTOR_PANEL_WIDTH = 320;
const MIN_WORKSPACE_PANEL_WIDTH = 400;
const SPLIT_RESIZE_HANDLE_WIDTH = 8;
const CHAT_PANEL_KEYBOARD_STEP = 16;
const DESIGN_SYSTEM_AUDIT_AUTO_REPAIR_ATTEMPTS = 2;
// Trailing-debounce window for the canonical (daemon + SQLite) tab-state write.
// Embedded-browser navigation bursts settle well within this; the local cache
// is written immediately so nothing is lost if the daemon write is coalesced.
const TAB_PERSIST_DEBOUNCE_MS = 400;
const MIN_NORMAL_SPLIT_WIDTH =
  MIN_CHAT_PANEL_WIDTH + SPLIT_RESIZE_HANDLE_WIDTH + MIN_WORKSPACE_PANEL_WIDTH;
type DesignSystemReviewEntry = NonNullable<ProjectMetadata['designSystemReview']>[string];
type DesignSystemReviewAgentTask = NonNullable<DesignSystemReviewEntry['agentTask']>;
interface DesignSystemReviewDetails {
  feedback?: string;
  files?: string[];
  agentTask?: DesignSystemReviewAgentTask;
}

function workspacePanelMinWidthForSplit(splitWidth: number): number {
  if (!Number.isFinite(splitWidth) || splitWidth <= 0) return MIN_WORKSPACE_PANEL_WIDTH;
  return splitWidth < MIN_NORMAL_SPLIT_WIDTH ? 0 : MIN_WORKSPACE_PANEL_WIDTH;
}

function maxChatPanelWidthForSplit(splitWidth: number): number {
  if (!Number.isFinite(splitWidth) || splitWidth <= 0) return MAX_CHAT_PANEL_WIDTH;
  const workspaceMinWidth = workspacePanelMinWidthForSplit(splitWidth);
  const viewportAwareMax = splitWidth - SPLIT_RESIZE_HANDLE_WIDTH - workspaceMinWidth;
  return Math.max(0, Math.min(MAX_CHAT_PANEL_WIDTH, Math.floor(viewportAwareMax)));
}

function clampPreferredChatPanelWidth(width: number): number {
  return Math.min(MAX_CHAT_PANEL_WIDTH, Math.max(MIN_CHAT_PANEL_WIDTH, Math.round(width)));
}

function clampChatPanelWidth(width: number, maxWidth = MAX_CHAT_PANEL_WIDTH): number {
  const effectiveMax = Math.max(0, Math.min(MAX_CHAT_PANEL_WIDTH, Math.floor(maxWidth)));
  const effectiveMin = Math.min(MIN_CHAT_PANEL_WIDTH, effectiveMax);
  return Math.min(effectiveMax, Math.max(effectiveMin, Math.round(width)));
}

function designSystemFeedbackAttachments(
  projectFiles: ProjectFile[],
  sectionFiles: string[],
): ChatAttachment[] {
  const fileLookup = new Map(projectFiles.map((file) => [file.name, file]));
  return sectionFiles
    .map((name) => fileLookup.get(name))
    .filter((file): file is ProjectFile => Boolean(file))
    .slice(0, 8)
    .map((file) => ({
      path: file.name,
      name: file.name,
      kind: file.kind === 'image' ? 'image' : 'file',
      size: file.size,
    }));
}

function chatAttachmentsFromPreviewCommentImages(
  images: PreviewCommentAttachment[] | undefined,
): ChatAttachment[] {
  if (!Array.isArray(images)) return [];
  const seen = new Set<string>();
  const out: ChatAttachment[] = [];
  for (const image of images) {
    const path = image.path.trim();
    if (!path || seen.has(path)) continue;
    seen.add(path);
    out.push({
      path,
      name: image.name.trim() || path.split('/').pop() || path,
      kind: 'image',
    });
  }
  return out;
}

function mergeChatAttachments(...groups: ChatAttachment[][]): ChatAttachment[] {
  const seen = new Set<string>();
  const out: ChatAttachment[] = [];
  for (const group of groups) {
    for (const attachment of group) {
      const path = attachment.path.trim();
      if (!path || seen.has(path)) continue;
      seen.add(path);
      out.push({ ...attachment, path });
    }
  }
  return out;
}

function historyWithWorkspaceContext(
  history: ChatMessage[],
  messageId: string,
  context: ChatSendMeta['context'] | undefined,
): ChatMessage[] {
  const items = context?.workspaceItems ?? [];
  if (items.length === 0) return history;
  const block = [
    '',
    '',
    '<active-workspace-context>',
    'Open Design selected the currently focused workspace tab as the default context for this turn.',
    ...items.map((item, index) => {
      const details = [
        item.path ? `path: ${item.path}` : null,
        item.absolutePath ? `absolute: ${item.absolutePath}` : null,
        item.url ? `url: ${item.url}` : null,
        item.title ? `title: ${item.title}` : null,
        item.tabId ? `tab: ${item.tabId}` : null,
      ].filter(Boolean).join(' | ');
      return `${index + 1}. ${item.kind}: ${item.label}${details ? ` | ${details}` : ''}`;
    }),
    '</active-workspace-context>',
  ].join('\n');
  return history.map((message) =>
    message.id === messageId && message.role === 'user'
      ? { ...message, content: `${message.content}${block}` }
      : message,
  );
}

function commentTaskQuery(attachment: ChatCommentAttachment): string {
  return (attachment.comment ?? '').trim();
}

function commentTaskContextAttachment(attachment: ChatCommentAttachment): ChatCommentAttachment {
  return {
    ...attachment,
    comment: '',
    commentContext: 'query',
  };
}

function designSystemNeedsWorkPrompt(
  sectionTitle: string,
  feedback: string,
  sectionFiles: string[],
): string {
  const fileList =
    sectionFiles.length > 0
      ? sectionFiles.map((name) => `- @${name}`).join('\n')
      : '- No generated files are registered for this section yet.';
  return (
    `Needs work on the design system section "${sectionTitle}".\n\n` +
    `User feedback:\n${feedback}\n\n` +
    `Relevant section files:\n${fileList}\n\n` +
    'Revise the design-system project files directly. Keep DESIGN.md, tokens, previews, UI kit examples, and assets consistent with the feedback. ' +
    'After editing, summarize what changed and which files should be reviewed again.'
  );
}

function readSavedChatPanelWidth(): number {
  if (typeof window === 'undefined') return DEFAULT_CHAT_PANEL_WIDTH;
  try {
    const raw = window.localStorage.getItem(CHAT_PANEL_WIDTH_STORAGE_KEY);
    const parsed = raw ? Number.parseInt(raw, 10) : Number.NaN;
    return Number.isFinite(parsed)
      ? clampPreferredChatPanelWidth(parsed)
      : DEFAULT_CHAT_PANEL_WIDTH;
  } catch {
    return DEFAULT_CHAT_PANEL_WIDTH;
  }
}

function saveChatPanelWidth(width: number): void {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(
      CHAT_PANEL_WIDTH_STORAGE_KEY,
      String(clampPreferredChatPanelWidth(width)),
    );
  } catch {
    // localStorage can be unavailable in hardened browser contexts.
  }
}

function autoSendFirstMessageKey(projectId: string): string {
  return `od:auto-send-first:${projectId}`;
}

function autoSendAttachmentsKey(projectId: string): string {
  return `od:auto-send-attachments:${projectId}`;
}

function designSystemAuditAutoRepairKey(projectId: string): string {
  return `od:design-system-audit-auto-repair:${projectId}`;
}

function readAutoSendAttachments(projectId: string): ChatAttachment[] {
  if (typeof window === 'undefined') return [];
  try {
    const raw = window.sessionStorage.getItem(autoSendAttachmentsKey(projectId));
    if (!raw) return [];
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isStoredChatAttachment);
  } catch {
    return [];
  }
}

function clearAutoSendSession(projectId: string): void {
  if (typeof window === 'undefined') return;
  try {
    window.sessionStorage.removeItem(autoSendFirstMessageKey(projectId));
    window.sessionStorage.removeItem(autoSendAttachmentsKey(projectId));
  } catch {
    /* ignore */
  }
}

function markDesignSystemAuditAutoRepairEligible(projectId: string): void {
  if (typeof window === 'undefined') return;
  try {
    window.sessionStorage.setItem(
      designSystemAuditAutoRepairKey(projectId),
      String(DESIGN_SYSTEM_AUDIT_AUTO_REPAIR_ATTEMPTS),
    );
  } catch {
    /* ignore */
  }
}

function consumeDesignSystemAuditAutoRepair(projectId: string): boolean {
  if (typeof window === 'undefined') return false;
  try {
    const key = designSystemAuditAutoRepairKey(projectId);
    const raw = window.sessionStorage.getItem(key);
    const attemptsRemaining = raw ? Number.parseInt(raw, 10) : 0;
    if (!Number.isFinite(attemptsRemaining) || attemptsRemaining <= 0) {
      window.sessionStorage.removeItem(key);
      return false;
    }
    const nextAttemptsRemaining = attemptsRemaining - 1;
    if (nextAttemptsRemaining > 0) {
      window.sessionStorage.setItem(key, String(nextAttemptsRemaining));
    } else {
      window.sessionStorage.removeItem(key);
    }
    return true;
  } catch {
    return false;
  }
}

function clearDesignSystemAuditAutoRepair(projectId: string): void {
  if (typeof window === 'undefined') return;
  try {
    window.sessionStorage.removeItem(designSystemAuditAutoRepairKey(projectId));
  } catch {
    /* ignore */
  }
}

function isDesignSystemWorkspaceMetadata(metadata: ProjectMetadata | undefined): boolean {
  return metadata?.importedFrom === 'design-system';
}

function isStoredChatAttachment(value: unknown): value is ChatAttachment {
  if (value === null || typeof value !== 'object') return false;
  const record = value as Record<string, unknown>;
  return (
    typeof record.path === 'string' &&
    record.path.length > 0 &&
    typeof record.name === 'string' &&
    record.name.length > 0 &&
    (record.kind === 'image' || record.kind === 'file') &&
    (record.size === undefined || typeof record.size === 'number') &&
    (record.order === undefined || typeof record.order === 'number')
  );
}

function workspaceContextItemEqual(
  a: WorkspaceContextItem | null,
  b: WorkspaceContextItem | null,
): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (
    a.id === b.id &&
    a.kind === b.kind &&
    a.label === b.label &&
    (a.tabId ?? '') === (b.tabId ?? '') &&
    (a.path ?? '') === (b.path ?? '') &&
    (a.absolutePath ?? '') === (b.absolutePath ?? '') &&
    (a.url ?? '') === (b.url ?? '') &&
    (a.title ?? '') === (b.title ?? '')
  );
}

function workspaceContextItemsEqual(
  a: WorkspaceContextItem[],
  b: WorkspaceContextItem[],
): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  return a.every((item, index) => workspaceContextItemEqual(item, b[index] ?? null));
}

function appendLiveArtifactEventItem(
  prev: LiveArtifactEventItem[],
  event: LiveArtifactEventItem['event'],
): LiveArtifactEventItem[] {
  liveArtifactEventSequence += 1;
  const next = [...prev, { id: liveArtifactEventSequence, event }];
  return next.length > 50 ? next.slice(next.length - 50) : next;
}

export function projectSplitClassName(workspaceFocused: boolean): string {
  return workspaceFocused ? 'split split-focus' : 'split';
}

// React key for the on-screen question form. Deliberately does NOT include the
// form's parsed `id`: there is at most one (first) form per assistant message,
// so `${conversation}:${message}` is already a stable, unique identity for the
// occurrence. Folding the parsed id in would remount the panel mid-stream — the
// preview shows the `discovery` fallback until the body `id` streams in, and a
// form that emits answerable questions before its `id` would flip identity
// while the user is mid-answer, dropping their selections. A distinct later
// form lives in a different assistant message, so it still gets its own key
// (and replays the reveal) without relying on the id.
export function buildQuestionFormKey(
  conversationId: string | null,
  assistantMessageId: string | null,
  hasForm: boolean,
): string | null {
  return conversationId && assistantMessageId && hasForm
    ? `${conversationId}:${assistantMessageId}`
    : null;
}

type ProjectSplitStyle = CSSProperties & {
  '--project-chat-panel-width': string;
  '--project-workspace-panel-track': string;
};

export function projectSplitStyle(
  workspaceFocused: boolean,
  chatPanelWidth: number,
  workspacePanelTrack: string,
): ProjectSplitStyle | undefined {
  if (workspaceFocused) return undefined;
  return {
    '--project-chat-panel-width': `${chatPanelWidth}px`,
    '--project-workspace-panel-track': workspacePanelTrack,
    gridTemplateColumns: `${chatPanelWidth}px ${SPLIT_RESIZE_HANDLE_WIDTH}px ${workspacePanelTrack}`,
  };
}

function applySplitChatPanelWidth(
  split: HTMLDivElement | null,
  width: number,
  workspacePanelTrack: string,
): void {
  if (!split) return;
  split.style.setProperty('--project-chat-panel-width', `${width}px`);
  split.style.gridTemplateColumns =
    `${width}px ${SPLIT_RESIZE_HANDLE_WIDTH}px ${workspacePanelTrack}`;
}

function shouldFetchElevenLabsVoiceOptions(project: Project): boolean {
  const metadata = project.metadata;
  return metadata?.kind === 'audio'
    && metadata.audioKind === 'speech'
    && metadata.audioModel === 'elevenlabs-v3'
    && !metadata.voice;
}

// The media model the user picked in the New Project → Media dialog, keyed by
// surface. For BYOK providers (AIHubMix) media is produced by the generate_*
// chat tools whose default model comes from the per-request byok*Model field —
// NOT the `od media generate` dispatcher — so without this seed the dialog pick
// is dropped and the conversation falls back to the Settings default. Returns
// undefined for non-media projects (and when the field is empty) so callers fall
// back to the Settings default exactly as before. The daemon re-validates the id
// against the active provider's registry, so a mismatched pick is safely ignored.
function projectMediaModelSeed(
  metadata: ProjectMetadata | null | undefined,
  surface: 'image' | 'video' | 'speech',
): string | undefined {
  if (!metadata) return undefined;
  if (surface === 'image' && metadata.kind === 'image') {
    return metadata.imageModel?.trim() || undefined;
  }
  if (surface === 'video' && metadata.kind === 'video') {
    return metadata.videoModel?.trim() || undefined;
  }
  if (surface === 'speech' && metadata.kind === 'audio' && metadata.audioKind === 'speech') {
    return metadata.audioModel?.trim() || undefined;
  }
  return undefined;
}

function projectMediaVoiceSeed(
  metadata: ProjectMetadata | null | undefined,
): string | undefined {
  if (metadata?.kind === 'audio' && metadata.audioKind === 'speech') {
    return metadata.voice?.trim() || undefined;
  }
  return undefined;
}

// Carry the creation-time model pick into the conversation ONLY when it belongs
// to the active BYOK provider. Guards against clobbering a user's Settings
// default with a model from a different provider — e.g. a SenseAudio user whose
// image project was created with the dialog's default `gpt-image-2` keeps their
// configured SenseAudio model instead of being forced to the registry default.
// AIHubMix's live (`aihubmix-` prefixed) ids resolve via mediaModelProviderId
// without waiting on the async catalogue, so the AIHubMix path still seeds.
function byokModelSeedForProtocol(
  metadata: ProjectMetadata | null | undefined,
  surface: 'image' | 'video' | 'speech',
  protocol: string | undefined,
): string | undefined {
  const picked = projectMediaModelSeed(metadata, surface);
  if (!picked) return undefined;
  return mediaModelProviderId(picked) === protocol ? picked : undefined;
}

function projectEventToAgentEvent(evt: ProjectEvent): LiveArtifactEventItem['event'] | null {
  if (evt.type === 'file-changed') return null;
  if (evt.type === 'conversation-created') return null;
  if (evt.type === 'live_artifact') {
    return {
      kind: 'live_artifact',
      action: evt.action,
      projectId: evt.projectId,
      artifactId: evt.artifactId,
      title: evt.title,
      refreshStatus: evt.refreshStatus,
    };
  }
  return {
    kind: 'live_artifact_refresh',
    phase: evt.phase,
    projectId: evt.projectId,
    artifactId: evt.artifactId,
    refreshId: evt.refreshId,
    title: evt.title,
    refreshedSourceCount: evt.refreshedSourceCount,
    error: evt.error,
  };
}

function artifactWithHtml(
  artifact: Artifact | null,
  fallbackIdentifier: string,
  html: string,
): Artifact {
  return artifact
    ? { ...artifact, html }
    : {
        identifier: fallbackIdentifier,
        title: '',
        html,
      };
}

export function ProjectView({
  project,
  routeFileName,
  routeConversationId = null,
  config,
  agents,
  skills,
  designTemplates,
  designSystems,
  daemonLive,
  onModeChange,
  onAgentChange,
  onAgentModelChange,
  onApiModelChange,
  onRefreshAgents,
  onThemeChange,
  onOpenSettings,
  onOpenAmrSettings,
  onOpenMcpSettings,
  onBrowsePlugins,
  onOpenConnectors,
  onAdoptPetInline,
  onTogglePet,
  onOpenPetSettings,
  onBack,
  onClearPendingPrompt,
  onTouchProject,
  onProjectChange,
  onProjectsRefresh,
  onChangeDefaultDesignSystem,
  onDesignSystemsRefresh,
}: Props) {
  const { locale, t } = useI18n();
  const analytics = useAnalytics();
  const iframeKeepAlivePool = useIframeKeepAlivePool();
  const handleThemeChange = onThemeChange ?? (() => {});
  // P0 page_view page_name=chat_panel — fire once per project mount.
  // ProjectView outlives conversation switches (ChatPane is keyed by
  // activeConversationId so it remounts when the user switches chats,
  // but this component does not), so page_view stays a "chat-panel
  // entry" metric instead of becoming a "conversation switch" count.
  // Reviewer #2285 (mrcfps, 2026-05-20 04:08) flagged the previous
  // ChatComposer-level emit for skewing the funnel.
  const chatPanelPageViewFiredRef = useRef<string | null>(null);
  const mountedRef = useRef(true);
  const trackedTimeoutsRef = useRef<Set<ReturnType<typeof setTimeout>>>(new Set());

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      for (const timer of trackedTimeoutsRef.current) clearTimeout(timer);
      trackedTimeoutsRef.current.clear();
    };
  }, []);

  const scheduleProjectTimeout = useCallback((callback: () => void, delayMs: number) => {
    if (!mountedRef.current) return null;
    const timer = setTimeout(() => {
      trackedTimeoutsRef.current.delete(timer);
      if (!mountedRef.current) return;
      callback();
    }, delayMs);
    trackedTimeoutsRef.current.add(timer);
    return timer;
  }, []);

  const clearProjectTimeout = useCallback((timer: ReturnType<typeof setTimeout> | null) => {
    if (timer == null) return;
    clearTimeout(timer);
    trackedTimeoutsRef.current.delete(timer);
  }, []);

  useEffect(() => {
    if (chatPanelPageViewFiredRef.current === project.id) return;
    chatPanelPageViewFiredRef.current = project.id;
    trackPageView(analytics.track, { page_name: 'chat_panel' });
    // 新手引导的第4步（"生成进度页"）在这里触发，而不是在
    // `DesignSystemDetailView` 中触发：生成路径直接导航到
    // 项目的 chat_panel（聊天面板），而不是设计系统详情页面。
    // 如果 sessionStorage 中仍存在新手引导的会话 ID，
    // 我们在这里标记漏斗的最后一行并清除它，
    // 这样后续任何对设计系统的访问都不会继承该归因。
    // E2E 测试（2026-05-21）确认这是用户实际执行的唯一路径 ——
    // 观察到：page_view chat_panel 事件会触发，但
    // page_view design_system_project 从未触发，
    // 因为从嵌入式新手引导的生成功能出发，不会访问那个路由。
    const onboardingSessionId = peekOnboardingSessionId();
    if (onboardingSessionId) {
      trackPageView(analytics.track, {
        page_name: 'onboarding',
        area: 'generation_progress',
        step_index: 'progress',
        step_name: 'generation',
        onboarding_session_id: onboardingSessionId,
      });
      clearOnboardingSessionId();
    }
  }, [analytics.track, project.id]);
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const conversationsRef = useRef<Conversation[]>([]);
  useEffect(() => {
    conversationsRef.current = conversations;
  }, [conversations]);
  const [activeConversationId, setActiveConversationId] = useState<string | null>(
    null,
  );
  const activeConversation = useMemo(
    () => conversations.find((conversation) => conversation.id === activeConversationId) ?? null,
    [conversations, activeConversationId],
  );
  const activeSessionMode = activeConversation?.sessionMode ?? 'design';
  const [messagesConversationId, setMessagesConversationId] = useState<string | null>(null);
  const [failedMessagesConversationId, setFailedMessagesConversationId] = useState<string | null>(null);
  const [conversationLoadError, setConversationLoadError] = useState<string | null>(null);
  const [messageLoadRetryNonce, setMessageLoadRetryNonce] = useState(0);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [forkingMessageId, setForkingMessageId] = useState<string | null>(null);
  const [activePluginActionPaths, setActivePluginActionPaths] = useState<Set<string>>(() => new Set());
  const [hiddenAssistantPluginActionPaths, setHiddenAssistantPluginActionPaths] = useState<Set<string>>(() => new Set());
  const [forceStreamingPluginMessageIds, setForceStreamingPluginMessageIds] = useState<Set<string>>(() => new Set());
  // Ephemeral, live-only accumulation of a tool call's streaming JSON input,
  // keyed by tool-use id (globally unique per run). Fed by `onToolInputDelta`
  // while the model is still emitting `input_json_delta`; dropped per-id once
  // the full `tool_use` lands and wiped when the run ends. Never persisted —
  // see daemon `daemonAgentPayloadToPersistedAgentEvent` (returns null).
  // `seq` records how many persisted events existed when the tool started
  // streaming, so the renderer can place the live card at the tool call's
  // position in the message (text before it = preamble, after it = hedging).
  const [liveToolInput, setLiveToolInput] = useState<Record<string, { name: string; text: string; seq: number }>>({});
  // True once the initial DB read for the active conversation has settled.
  // Auto-send gates on this so it can't fire before listMessages resolves and
  // race-clobber the freshly-pushed user + assistant placeholder. Without
  // this, the auto-send writes [user, assistant] into state, then the still
  // in-flight listMessages PUT response arrives, runs setMessages(list), and
  // wipes both — leaving the daemon's run with no client-side message to
  // attach the runId to.
  const [messagesInitialized, setMessagesInitialized] = useState(false);
  const [previewComments, setPreviewComments] = useState<PreviewComment[]>([]);
  // Mirror so the send-now interrupt path can read the current statuses
  // synchronously without re-creating its callback on every comment change.
  const previewCommentsRef = useRef<PreviewComment[]>([]);
  useEffect(() => {
    previewCommentsRef.current = previewComments;
  }, [previewComments]);
  const [attachedComments, setAttachedComments] = useState<PreviewComment[]>([]);
  const [streaming, setStreaming] = useState(false);
  const [streamingConversationId, setStreamingConversationId] = useState<string | null>(null);
  // Safety net: drop any live tool-input partials whose tool never produced a
  // full `tool_use` (run errored/canceled mid-call) once streaming settles.
  useEffect(() => {
    if (!streaming) setLiveToolInput((prev) => (Object.keys(prev).length ? {} : prev));
  }, [streaming]);
  const [error, setError] = useState<string | null>(null);
  const [audioVoiceOptionsError, setAudioVoiceOptionsError] = useState<string | null>(null);
  const [artifact, setArtifact] = useState<Artifact | null>(null);
  const [filesRefresh, setFilesRefresh] = useState(0);
  // True while a working-dir replace is reindexing the new folder. Surfaced
  // to the Design Files panel so the file list shows a loading state instead
  // of silently sitting on the old tree for the few seconds the scan takes.
  const [projectFiles, setProjectFiles] = useState<ProjectFile[]>([]);
  const projectFilesRef = useRef<ProjectFile[]>([]);
  const [liveArtifacts, setLiveArtifacts] = useState<LiveArtifactSummary[]>([]);
  const [liveArtifactEvents, setLiveArtifactEvents] = useState<LiveArtifactEventItem[]>([]);
  const [workspaceFocused, setWorkspaceFocused] = useState(false);
  const [commentInspectorActive, setCommentInspectorActive] = useState(false);
  const commentInspectorPortalId = useId();
  const leftInspectorActive = commentInspectorActive;
  // Per-session override for the BYOK chat's generate_image tool. Seeded once
  // from the New Project → Media model pick (project.metadata.imageModel) — but
  // only when that pick belongs to the active BYOK provider (see
  // byokModelSeedForProtocol) — falling back to the Settings default
  // (config.byokImageModel) otherwise. Subsequent selections live only in this
  // component's state — page refresh / project switch resets to this seed.
  // Persistent defaults live in Settings → BYOK → Image generation model.
  const [byokImageModelOverride, setByokImageModelOverride] = useState<string>(
    () => byokModelSeedForProtocol(project.metadata, 'image', config.apiProtocol) ?? config.byokImageModel ?? '',
  );
  // Same per-session override for the BYOK chat's generate_video tool, seeded
  // from the project's videoModel pick (provider-gated), then Settings.
  const [byokVideoModelOverride, setByokVideoModelOverride] = useState<string>(
    () => byokModelSeedForProtocol(project.metadata, 'video', config.apiProtocol) ?? config.byokVideoModel ?? '',
  );
  // Same per-session overrides for the BYOK chat's generate_speech tool (model +
  // voice), seeded from the project's speech pick (provider-gated), then Settings.
  const [byokSpeechModelOverride, setByokSpeechModelOverride] = useState<string>(
    () => byokModelSeedForProtocol(project.metadata, 'speech', config.apiProtocol) ?? config.byokSpeechModel ?? '',
  );
  // Voice only carries when the speech model itself is carried (same provider),
  // so a cross-provider voice id never leaks into the request.
  const [byokSpeechVoiceOverride, setByokSpeechVoiceOverride] = useState<string>(
    () => (byokModelSeedForProtocol(project.metadata, 'speech', config.apiProtocol)
      ? projectMediaVoiceSeed(project.metadata)
      : undefined) ?? config.byokSpeechVoice ?? '',
  );
  // Live model option lists (same hooks the composer/Settings pickers use) so
  // the chat "default" (no explicit pick) resolves to the FIRST catalogue model
  // shown in the dropdown — not a hardcoded id. The daemon keeps its own
  // fallback for when the catalogue hasn't loaded.
  const byokImageModelOptionsPV = useByokImageModelOptions(config.apiProtocol);
  const byokVideoModelOptionsPV = useByokVideoModelOptions(config.apiProtocol);
  const byokSpeechModelOptionsPV = useByokSpeechModelOptions(config.apiProtocol);
  // PR #974 round 7 (mrcfps @ useDesignMdState.ts:131): counter that
  // bumps on file-changed SSE events, live_artifact* events, and the
  // chat streaming-completion edge so the staleness chip stays in sync
  // with the underlying mtimes / conversation updatedAt as the user
  // keeps working post-finalize. The hook treats it as a dep and
  // recomputes whenever it changes.
  const [designMdRefreshKey, setDesignMdRefreshKey] = useState(0);
  // ----- Continue in CLI / Finalize design package wiring (#451) -----
  // The toast surface is shared between Finalize errors and the
  // success/fallback toasts emitted from handleContinueInCli.
  const projectDetail = useProjectDetail(project.id);
  const designMdState = useDesignMdState(project.id, designMdRefreshKey);
  const finalize = useFinalizeProject(project.id);
  const terminalLauncher = useTerminalLaunch();
  const [projectActionsToast, setProjectActionsToast] = useState<{
    message: string;
    details: string | null;
    code?: string | null;
  } | null>(null);
  const [chatSeed, setChatSeed] = useState<{ id: string; value: string } | null>(null);
  const [autoAuditRepairSeed, setAutoAuditRepairSeed] =
    useState<{ id: string; value: string } | null>(null);
  const [chatPanelWidth, setChatPanelWidth] = useState(readSavedChatPanelWidth);
  const [chatPanelMaxWidth, setChatPanelMaxWidth] = useState(MAX_CHAT_PANEL_WIDTH);
  const [workspacePanelMinWidth, setWorkspacePanelMinWidth] = useState(MIN_WORKSPACE_PANEL_WIDTH);
  const [resizingChatPanel, setResizingChatPanel] = useState(false);
  const splitRef = useRef<HTMLDivElement | null>(null);
  const chatPanelWidthRef = useRef(chatPanelWidth);
  const preferredChatPanelWidthRef = useRef(chatPanelWidth);
  const resizeStartPreferredWidthRef = useRef(chatPanelWidth);
  const chatPanelMaxWidthRef = useRef(chatPanelMaxWidth);
  const resizeStateRef = useRef<{
    startClientX: number;
    startWidth: number;
    isRtl: boolean;
    hasMoved: boolean;
  } | null>(null);
  const pointerCleanupRef = useRef<(() => void) | null>(null);
  const pointerFrameRef = useRef<number | null>(null);
  const pendingPointerClientXRef = useRef<number | null>(null);
  // The persisted set of open tabs + active tab. Persisted via PUT on every
  // change; loaded once when the project mounts.
  const [openTabsState, setOpenTabsState] = useState<OpenTabsState>({
    tabs: [],
    active: null,
  });
  // Artifact context for the header actions (settings gear, handoff) that live
  // in this workspace's header alongside FileViewer's present/share/download.
  // Mirrors the artifact_id / artifact_kind that FileViewer attaches, derived
  // from the currently-active file tab, so all artifact_header analytics carry
  // the same dimensions. Undefined on non-file tabs (e.g. the file list).
  const headerArtifact = useMemo<{
    artifact_id?: string;
    artifact_kind?: TrackingArtifactKind;
  }>(() => {
    const activeName = openTabsState.active;
    const file = activeName
      ? projectFiles.find((entry) => entry.name === activeName) ?? null
      : null;
    if (!file) return {};
    return {
      artifact_id: anonymizeArtifactId({ projectId: project.id, fileName: file.name }),
      artifact_kind: artifactKindToTracking({ fileKind: file.kind ?? null }),
    };
  }, [openTabsState.active, projectFiles, project.id]);
  const routeFileNameRef = useRef(routeFileName);
  routeFileNameRef.current = routeFileName;
  const [activeWorkspaceContext, setActiveWorkspaceContext] =
    useState<WorkspaceContextItem | null>(null);
  const [workspaceContexts, setWorkspaceContexts] = useState<WorkspaceContextItem[]>([]);
  const tabsLoadedRef = useRef(false);
  const tabsHydratedFromSavedStateRef = useRef(false);
  const hasAppliedInitialPrimaryOpenRef = useRef(false);
  // Routed to FileWorkspace — bumped whenever the user clicks "open" on a
  // tool card, an attachment chip, or a produced-file chip in chat. We
  // include a nonce so re-clicking the same name after the user closed the
  // tab still focuses it.
  const [openRequest, setOpenRequest] = useState<{ name: string; nonce: number } | null>(null);
  // Like `openRequest`, but additionally asks the preview workspace to open the
  // file's Share/Export menu. Drives the "Share" next-step action: it reuses the
  // existing export/deploy surface rather than introducing a new share backend.
  const [shareRequest, setShareRequest] = useState<{ name: string; nonce: number } | null>(null);
  // Parallel to shareRequest, but opens the workspace's Download/Export menu.
  const [downloadRequest, setDownloadRequest] = useState<{ name: string; nonce: number } | null>(null);
  // When a queued chat send starts processing, ask the workspace to flip the
  // deck preview to the slide its marked element lives on, so the user watches
  // the edit land in context instead of staying parked on slide 1. Mirrors the
  // `shareRequest` nonce signal: FileWorkspace matches `name` against the open
  // file and FileViewer consumes each nonce once.
  const [slideNavRequest, setSlideNavRequest] = useState<
    { name: string; slideIndex: number; nonce: number } | null
  >(null);
  const abortRef = useRef<AbortController | null>(null);
  const cancelRef = useRef<AbortController | null>(null);
  // Runs explicitly superseded by a "send now" interrupt. Their abort
  // controller is recorded here synchronously — before handleStop() clears the
  // active refs — so the run's late terminal callbacks (which the daemon still
  // delivers for a canceled run) can be recognized as stale and skip every
  // current-run side effect, independent of abortRef churn. A WeakSet so a
  // finished run's controller is collected once nothing else references it.
  const supersededRunsRef = useRef<WeakSet<AbortController>>(new WeakSet());
  const streamingConversationIdRef = useRef<string | null>(null);
  const [queuedChatSends, setQueuedChatSends] = useState<QueuedChatSend[]>([]);
  const queuedChatSendsRef = useRef<QueuedChatSend[]>([]);
  const sendTextBufferRef = useRef<BufferedTextUpdates | null>(null);
  const reattachTextBuffersRef = useRef<Set<BufferedTextUpdates>>(new Set());
  const reattachControllersRef = useRef<Map<string, AbortController>>(new Map());
  const reattachCancelControllersRef = useRef<Map<string, AbortController>>(new Map());
  const completedReattachRunsRef = useRef<Set<string>>(new Set());
  const recoveredArtifactMessagesRef = useRef<Set<string>>(new Set());
  const messagesRef = useRef<ChatMessage[]>([]);
  const startingQueuedChatSendIdRef = useRef<string | null>(null);
  const [queuedAutoStartTick, setQueuedAutoStartTick] = useState(0);
  const skillCache = useRef<Map<string, string>>(new Map());
  const designCache = useRef<Map<string, string>>(new Map());
  const templateCache = useRef<Map<string, ProjectTemplate>>(new Map());
  // We auto-save the most recent artifact to the project folder. Track the
  // last name we persisted so re-renders during streaming don't spawn
  // duplicate writes.
  const savedArtifactRef = useRef<string | null>(null);
  // Pending Write tool invocations: tool_use_id -> destination basename.
  // When the matching tool_result lands we refresh the file list and open
  // the file as a tab once. Keying off the tool_use_id (rather than
  // diffing the file list at end-of-turn) lets us auto-open the moment
  // the agent's Write actually completes, without the previous synthetic
  // "live" tab that was causing flicker against manual opens.
  const pendingWritesRef = useRef<Map<string, string>>(new Map());
  // Track which conversation the current messages belong to, so we can
  // correctly gate new-conversation creation even during async loads.
  const messagesConversationIdRef = useRef<string | null>(null);
  const creatingConversationRef = useRef(false);
  // Last conversation id this view pushed into the URL. Lets the
  // route -> active-conversation sync tell a genuine external navigation
  // apart from the URL merely lagging a local conversation switch.
  const lastSyncedConversationIdRef = useRef<string | null>(null);
  // Live mirror of the currently-viewed project id. Used to bail out of
  // the conversation-created async refresh (#1361) if the user switches
  // projects while the refetch is in flight — the existing project-load
  // effects use the same kind of cancellation guard.
  const projectIdRef = useRef(project.id);
  useEffect(() => {
    projectIdRef.current = project.id;
  }, [project.id]);
  useEffect(() => {
    messagesRef.current = messages;
  }, [messages]);
  useEffect(() => {
    setChatSeed(null);
    setAutoAuditRepairSeed(null);
    const restored = loadQueuedChatSends(project.id);
    queuedChatSendsRef.current = restored;
    setQueuedChatSends(restored);
  }, [project.id]);
  // Monotonic token bumped on every `conversation-created` refresh dispatch.
  // Two rapid events (e.g. concurrent routine runs against the same reused
  // project, #1502) can start overlapping `listConversations` calls; if the
  // later request resolves first with N+1 conversations and the earlier
  // request resolves afterwards with only N, an unconditional
  // `setConversations(list)` would drop the newest conversation. Each
  // dispatch captures the token at start; only the dispatch whose token
  // still equals `conversationsRefreshTokenRef.current` at await-return is
  // allowed to apply its result.
  const conversationsRefreshTokenRef = useRef(0);
  const [creatingConversation, setCreatingConversation] = useState(false);
  const currentConversationHasActiveRun = useMemo(
    () => messages.some((m) => m.role === 'assistant' && isActiveRunStatus(m.runStatus)),
    [messages],
  );
  const currentConversationHasRecoverableArtifact = useMemo(
    () => messages.some((message) => hasRecoverableArtifactMessage(message)),
    [messages],
  );
  const currentConversationLoading = Boolean(
    activeConversationId
      && messagesConversationId !== activeConversationId
      && failedMessagesConversationId !== activeConversationId,
  );
  const currentConversationStreaming = streaming && streamingConversationId === activeConversationId;
  const currentConversationBusy = currentConversationLoading
    || currentConversationStreaming
    || currentConversationHasActiveRun;
  const currentConversationAwaitingActiveRunAttach =
    currentConversationHasActiveRun && !currentConversationStreaming;
  const currentConversationSendDisabled = currentConversationLoading
    || failedMessagesConversationId === activeConversationId
    || currentConversationAwaitingActiveRunAttach;
  const currentConversationActionDisabled = currentConversationBusy || currentConversationSendDisabled;
  const currentConversationQueueDisabled = currentConversationLoading
    || failedMessagesConversationId === activeConversationId;

  // The discovery question form lives in the right-hand Questions tab. We
  // derive it from the latest assistant message: if that message embeds a
  // <question-form> block, the panel renders it. The form is interactive
  // only while it's the most recent turn and the user hasn't answered yet
  // (an answer arrives as a following "[form answers …]" user message).
  const lastAssistantIndex = useMemo(() => {
    for (let i = messages.length - 1; i >= 0; i--) {
      if (messages[i]?.role === 'assistant') return i;
    }
    return -1;
  }, [messages]);
  const lastAssistantContent =
    lastAssistantIndex >= 0 ? messages[lastAssistantIndex]?.content ?? '' : '';
  const lastAssistantMessageId =
    lastAssistantIndex >= 0 ? messages[lastAssistantIndex]?.id ?? null : null;
  const questionForm: QuestionForm | null = useMemo(
    () => findFirstQuestionForm(lastAssistantContent)?.form ?? null,
    [lastAssistantContent],
  );
  const questionFormSubmittedAnswers = useMemo(() => {
    if (!questionForm) return undefined;
    for (let i = lastAssistantIndex + 1; i < messages.length; i++) {
      const m = messages[i];
      if (m?.role !== 'user') continue;
      const parsed = parseSubmittedAnswers(questionForm, m.content ?? '');
      if (parsed) return parsed;
    }
    return undefined;
  }, [questionForm, lastAssistantIndex, messages]);
  const questionsGenerating =
    currentConversationStreaming && hasUnterminatedQuestionForm(lastAssistantContent);
  // While the form is still streaming, parse it tolerantly so the Questions tab
  // can show a frame (title) immediately and fill questions in as they arrive.
  const questionFormPreview = useMemo(
    () => (questionsGenerating ? parsePartialQuestionForm(lastAssistantContent) : null),
    [questionsGenerating, lastAssistantContent],
  );
  // The active (latest, unanswered) form stays editable the whole time it's on
  // screen — while it streams in AND while the turn is still busy — so it never
  // flickers between the locked (grey) and interactive (accent) styles.
  // Submission is gated separately by the panel via `submitDisabled`/generating.
  const questionFormActive =
    (!!questionForm || questionsGenerating) && questionFormSubmittedAnswers === undefined;
  // Mirror `questionFormActive`'s unanswered gate: once the user answers, the
  // Questions tab closes, so the auto-focus nonce must not treat an answered
  // form as a freshly appeared one.
  const hasQuestions =
    Boolean(questionForm || questionsGenerating) && questionFormSubmittedAnswers === undefined;
  // Stable identity for the current form occurrence, used to remember that its
  // one-by-one reveal already played. Keyed on the conversation + the hosting
  // assistant message id (not the message index, and NOT the parsed form id —
  // see buildQuestionFormKey). The assistant message id is allocated once and
  // kept in place across the streaming→persisted swap (same `assistantId`
  // throughout), so it survives the brief unmount/re-focus of the Questions tab
  // without replaying the animation, yet differs for every distinct form
  // occurrence (each lives in its own assistant message).
  const questionFormKey = useMemo(
    () =>
      buildQuestionFormKey(
        activeConversationId,
        lastAssistantMessageId,
        Boolean(questionForm ?? questionFormPreview),
      ),
    [activeConversationId, lastAssistantMessageId, questionForm, questionFormPreview],
  );

  // Release #3661: let a past question form be manually re-opened in the
  // Questions panel. Layered on top of main's stable questionFormKey (#3644) —
  // the `displayed*` values fall back to the live form when nothing is manually
  // pinned, so both fixes coexist.
  const [manualQuestionFormRequest, setManualQuestionFormRequest] =
    useState<QuestionFormOpenRequest | null>(null);
  useEffect(() => {
    setManualQuestionFormRequest(null);
  }, [project.id, activeConversationId]);
  useEffect(() => {
    if (hasQuestions && questionFormKey) setManualQuestionFormRequest(null);
  }, [hasQuestions, questionFormKey]);
  const displayedQuestionForm = manualQuestionFormRequest?.form ?? questionForm;
  const displayedQuestionFormPreview = manualQuestionFormRequest ? null : questionFormPreview;
  const displayedQuestionFormSubmittedAnswers =
    manualQuestionFormRequest?.submittedAnswers ?? questionFormSubmittedAnswers;
  const displayedQuestionFormActive = manualQuestionFormRequest ? false : questionFormActive;
  const displayedQuestionsGenerating = manualQuestionFormRequest ? false : questionsGenerating;
  const displayedQuestionFormKey = manualQuestionFormRequest
    ? `${activeConversationId ?? 'conversation'}:${manualQuestionFormRequest.messageId}:${manualQuestionFormRequest.form.id}:manual`
    : questionFormKey;

  // Auto-switch the workspace to the Questions tab when a new discovery form
  // first appears, and let the chat banner re-focus it on click. The nonce
  // bump is what FileWorkspace listens to.
  const [questionsFocusNonce, setQuestionsFocusNonce] = useState(0);
  const prevHasQuestionsRef = useRef(false);
  useEffect(() => {
    if (hasQuestions && !prevHasQuestionsRef.current) {
      setQuestionsFocusNonce((n) => n + 1);
    }
    prevHasQuestionsRef.current = hasQuestions;
  }, [hasQuestions]);
  const focusQuestionsRequest = useMemo(
    () => (questionsFocusNonce > 0 ? { nonce: questionsFocusNonce } : null),
    [questionsFocusNonce],
  );
  const submittedAnswersForQuestionFormRequest = useCallback((request: QuestionFormOpenRequest) => {
    const assistantIndex = messages.findIndex((m) => m.id === request.messageId);
    if (assistantIndex < 0) return null;
    for (let i = assistantIndex + 1; i < messages.length; i++) {
      const m = messages[i];
      if (!m) continue;
      if (m.role === 'assistant') break;
      if (m.role !== 'user') continue;
      const parsed = parseSubmittedAnswers(request.form, m.content ?? '');
      if (parsed) return parsed;
    }
    return null;
  }, [messages]);
  const openQuestionsTab = useCallback((request?: QuestionFormOpenRequest) => {
    if (request) {
      const opensCurrentLiveForm =
        request.messageId === lastAssistantMessageId
        && questionForm?.id === request.form.id
        && questionFormSubmittedAnswers === undefined;
      if (opensCurrentLiveForm) {
        setManualQuestionFormRequest(null);
      } else {
        setManualQuestionFormRequest({
          ...request,
          submittedAnswers:
            request.submittedAnswers ?? submittedAnswersForQuestionFormRequest(request) ?? undefined,
        });
      }
    }
    setQuestionsFocusNonce((n) => n + 1);
  }, [
    lastAssistantMessageId,
    questionForm,
    questionFormSubmittedAnswers,
    submittedAnswersForQuestionFormRequest,
  ]);

  const currentConversationQueuedItems = activeConversationId
    ? queuedChatSends
        .filter((item) => item.conversationId === activeConversationId)
        .map((item) => {
          const queuedItem = {
            id: item.id,
            prompt: item.prompt,
            attachments: item.attachments,
            commentAttachments: item.commentAttachments,
          };
          if (item.meta === undefined) return queuedItem;
          return { ...queuedItem, meta: item.meta };
        })
    : [];
  const newConversationDisabled = creatingConversation;
  const activeCompletionNotificationRunsRef = useRef<Set<string>>(new Set());
  const completedNotificationRunsRef = useRef<Set<string>>(new Set());

  // Load conversations on project switch. If none exist (older projects
  // pre-conversations, or a freshly created one whose default seed got
  // dropped), create one on the fly.
  useEffect(() => {
    let cancelled = false;
    setConversations([]);
    setActiveConversationId(null);
    setMessagesConversationId(null);
    setFailedMessagesConversationId(null);
    setMessageLoadRetryNonce(0);
    setConversationLoadError(null);
    setMessages([]);
    setPreviewComments([]);
    setAttachedComments([]);
    setStreaming(false);
    streamingConversationIdRef.current = null;
    setStreamingConversationId(null);
    setError(null);
    setAudioVoiceOptionsError(null);
    setArtifact(null);
    savedArtifactRef.current = null;
    pendingWritesRef.current.clear();
    (async () => {
      try {
        const list = await listConversations(project.id);
        if (cancelled) return;
        if (list.length === 0) {
          const fresh = await createConversation(project.id);
          if (cancelled) return;
          if (fresh) {
            setConversations([fresh]);
            setActiveConversationId(fresh.id);
          } else {
            throw new Error('Could not create a conversation for this project.');
          }
        } else {
          setConversations(list);
          // Issue #1505: when the URL deep-links to a specific
          // conversation, prefer that one. Falls through to list[0]
          // when the routed id is null or no longer present (the
          // routine row may have been deleted between the route
          // landing and the conversation list loading).
          const routedMatch = routeConversationId
            ? list.find((c) => c.id === routeConversationId) ?? null
            : null;
          setActiveConversationId(routedMatch ? routedMatch.id : list[0]!.id);
        }
      } catch (err) {
        if (cancelled) return;
        const message = err instanceof Error ? err.message : 'Could not load conversations for this project.';
        setConversations([]);
        setActiveConversationId(null);
        setConversationLoadError(message);
        setError(message);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [project.id]);

  // Issue #1505: when the URL changes the routed conversation id while
  // we are already inside the project (e.g. the user clicks "Open
  // project" on a different routine history row in the same project),
  // switch the active conversation without re-fetching the list.
  // Guards: only acts when the routed id is non-null AND present in
  // the already-loaded list, and only when it differs from the current
  // active id. Falls through to a no-op for stale / missing routes so
  // the default picker above keeps its result.
  useEffect(() => {
    if (!routeConversationId) {
      lastSeenRouteConversationIdRef.current = null;
      return;
    }
    if (conversations.length === 0) return;
    if (routeConversationId === activeConversationId) return;
    // When the route still points at the conversation this view last
    // pushed to the URL, the mismatch means a local switch (new
    // conversation, history pick) moved activeConversationId ahead and
    // the URL sync below has not caught up yet. Following the stale
    // route here would fight that sync and remount ChatPane in a loop,
    // so only react to a genuinely external navigation.
    if (routeConversationId === lastSyncedConversationIdRef.current) return;
    if (lastSeenRouteConversationIdRef.current === routeConversationId) return;
    lastSeenRouteConversationIdRef.current = routeConversationId;
    const match = conversations.find((c) => c.id === routeConversationId);
    if (!match) return;
    setActiveConversationId(routeConversationId);
  }, [routeConversationId, conversations, activeConversationId]);

  useEffect(() => {
    setWorkspaceFocused(false);
  }, [project.id]);

  // Load messages whenever the active conversation changes. This happens
  // on project mount (after conversations load) and on user-triggered
  // conversation switches.
  useEffect(() => {
    if (!activeConversationId) {
      setMessages([]);
      setMessagesInitialized(false);
      setPreviewComments([]);
      setAttachedComments([]);
      setMessagesConversationId(null);
      setFailedMessagesConversationId(null);
      messagesConversationIdRef.current = null;
      setStreaming(false);
      streamingConversationIdRef.current = null;
      setStreamingConversationId(null);
      return;
    }
    // Reset the initialized flag so auto-send waits for the new
    // conversation's DB read to settle before checking messages.length.
    setMessagesInitialized(false);
    let cancelled = false;
    setMessages([]);
    setPreviewComments([]);
    setAttachedComments([]);
    setArtifact(null);
    setMessagesConversationId(null);
    setFailedMessagesConversationId(null);
    setStreaming(false);
    streamingConversationIdRef.current = null;
    setStreamingConversationId(null);
    savedArtifactRef.current = null;
    pendingWritesRef.current.clear();
    if (messagesConversationIdRef.current !== activeConversationId) {
      messagesConversationIdRef.current = null;
    }
    (async () => {
      try {
        const [list, comments] = await Promise.all([
          listMessages(project.id, activeConversationId),
          fetchPreviewComments(project.id, activeConversationId),
        ]);
        if (cancelled) return;
        setMessages(list);
        setMessagesInitialized(true);
        setPreviewComments(comments);
        setAttachedComments([]);
        setArtifact(null);
        setError(null);
        savedArtifactRef.current = null;
        pendingWritesRef.current.clear();
        messagesConversationIdRef.current = activeConversationId;
        setMessagesConversationId(activeConversationId);
        setFailedMessagesConversationId(null);
      } catch (err) {
        if (cancelled) return;
        const message = err instanceof Error ? err.message : 'Could not load messages for this conversation.';
        setMessages([]);
        setPreviewComments([]);
        setAttachedComments([]);
        setArtifact(null);
        setError(message);
        savedArtifactRef.current = null;
        pendingWritesRef.current.clear();
        messagesConversationIdRef.current = null;
        setMessagesConversationId(null);
        setFailedMessagesConversationId(activeConversationId);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [project.id, activeConversationId, messageLoadRetryNonce]);

  useEffect(() => {
    return () => {
      sendTextBufferRef.current?.cancel();
      sendTextBufferRef.current = null;
      // Unmounts / conversation switches should only detach local stream
      // consumers. Aborting the daemon cancel controllers here turns routine
      // cleanup into an explicit POST /api/runs/:id/cancel, which can mark a
      // live run canceled even when the user never clicked Stop.
      abortRef.current?.abort();
      abortRef.current = null;
      cancelRef.current = null;
      for (const textBuffer of reattachTextBuffersRef.current) textBuffer.cancel();
      reattachTextBuffersRef.current.clear();
      for (const controller of reattachControllersRef.current.values()) {
        if (abortRef.current === controller) abortRef.current = null;
        controller.abort();
      }
      for (const controller of reattachCancelControllersRef.current.values()) {
        // Route changes should only detach the browser-side SSE listener.
        // Aborting this signal maps to POST /cancel, so leave the daemon run alive.
        if (cancelRef.current === controller) cancelRef.current = null;
      }
      reattachControllersRef.current.clear();
      reattachCancelControllersRef.current.clear();
    };
  }, [project.id, activeConversationId]);

  const cancelSendTextBuffer = useCallback((flushPending = false) => {
    if (flushPending) sendTextBufferRef.current?.flush();
    sendTextBufferRef.current?.cancel();
    sendTextBufferRef.current = null;
  }, []);

  const cancelReattachTextBuffers = useCallback((flushPending = false) => {
    for (const textBuffer of reattachTextBuffersRef.current) {
      if (flushPending) textBuffer.flush();
      textBuffer.cancel();
    }
    reattachTextBuffersRef.current.clear();
  }, []);

  const notifyCompletedRun = useCallback((last: ChatMessage) => {
    // Round 7 (mrcfps @ useDesignMdState.ts:131): a chat turn just
    // settled — conversation updatedAt almost certainly moved, so
    // recompute DESIGN.md staleness even when the turn produced no
    // file mutations or live artifacts.
    setDesignMdRefreshKey((n) => n + 1);

    const status = last.runStatus;
    if (status !== 'succeeded' && status !== 'failed') return;

    const cfg = config.notifications ?? DEFAULT_NOTIFICATIONS;
    if (cfg.soundEnabled) {
      playSound(status === 'succeeded' ? cfg.successSoundId : cfg.failureSoundId);
    }

    if (cfg.desktopEnabled) {
      // Successes only interrupt when the user is on another tab/window.
      // Failures alert regardless — losing a long agent run silently is
      // worse than a small interruption when the page is in focus.
      const isHidden = typeof document !== 'undefined' && document.hidden;
      const isFocused = typeof document === 'undefined' ? true : document.hasFocus();
      if (status === 'failed' || isHidden || !isFocused) {
        const title = status === 'succeeded'
          ? t('notify.successTitle')
          : t('notify.failureTitle');
        const fallbackBody = status === 'succeeded'
          ? t('notify.successBody')
          : t('notify.failureBody');
        const trimmed = (last.content ?? '').trim();
        const body = trimmed ? trimmed.slice(0, 80) : fallbackBody;
        void showCompletionNotification({
          status,
          title,
          body,
          onClick: () => {
            if (typeof window !== 'undefined') window.focus();
          },
        });
      }
    }
  }, [config.notifications, t]);

  // Fire completion feedback from assistant run-status transitions rather than
  // from the local SSE listener state. A run can finish while its conversation
  // is detached; when the user returns, the terminal status should still produce
  // the one completion notification for runs this view previously saw active.
  useEffect(() => {
    const completedMessages: ChatMessage[] = [];
    for (const message of messages) {
      if (message.role !== 'assistant') continue;
      const keys = message.runId ? [message.runId, message.id] : [message.id];
      if (isActiveRunStatus(message.runStatus)) {
        for (const key of keys) activeCompletionNotificationRunsRef.current.add(key);
        continue;
      }
      if (message.runStatus !== 'succeeded' && message.runStatus !== 'failed') continue;
      if (!keys.some((key) => activeCompletionNotificationRunsRef.current.has(key))) continue;
      if (keys.some((key) => completedNotificationRunsRef.current.has(key))) continue;
      for (const key of keys) completedNotificationRunsRef.current.add(key);
      completedMessages.push(message);
    }

    for (const message of completedMessages) notifyCompletedRun(message);
  }, [messages, notifyCompletedRun]);

  // Hydrate the open-tabs state once per project. After this initial
  // load, every mutation flows through saveTabsState() which keeps DB +
  // local state coherent.
  useEffect(() => {
    let cancelled = false;
    tabsLoadedRef.current = false;
    tabsHydratedFromSavedStateRef.current = false;
    hasAppliedInitialPrimaryOpenRef.current = false;
    setOpenTabsState({ tabs: [], active: null });
    (async () => {
      const state = await loadTabs(project.id);
      if (cancelled) return;
      const routeActive = routeFileNameRef.current;
      let nextState = routeActive
        ? {
            ...state,
            tabs: state.tabs.includes(routeActive)
              ? state.tabs
              : [...state.tabs, routeActive],
            active: routeActive,
          }
        : state;
      if (routeActive) {
        nextState = cacheTabsLocally(project.id, nextState);
        void persistTabsToDaemonNow(project.id, nextState);
      }
      tabsHydratedFromSavedStateRef.current = state.hasSavedState === true;
      setOpenTabsState(nextState);
      tabsLoadedRef.current = true;
    })();
    return () => {
      cancelled = true;
    };
  }, [project.id]);

  // Debounce the canonical (daemon + SQLite) tab-state write. The embedded
  // browser fans out url/title/favicon updates in bursts on a single page load
  // (did-navigate, did-navigate-in-page, page-title-updated, favicon), and each
  // used to be a localStorage write + HTTP PUT + SQLite UPDATE + re-render.
  // We keep React state and the local cache IMMEDIATE (so the UI and a reload
  // are never stale) and coalesce only the daemon PUT.
  const tabsDaemonSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const pendingDaemonTabsRef = useRef<OpenTabsState | null>(null);
  const flushTabsDaemonSave = useCallback(() => {
    if (tabsDaemonSaveTimerRef.current != null) {
      clearTimeout(tabsDaemonSaveTimerRef.current);
      tabsDaemonSaveTimerRef.current = null;
    }
    const pending = pendingDaemonTabsRef.current;
    pendingDaemonTabsRef.current = null;
    if (pending) void persistTabsToDaemonNow(project.id, pending);
  }, [project.id]);

  const persistTabsState = useCallback(
    (next: OpenTabsState) => {
      setOpenTabsState(next);
      if (!tabsLoadedRef.current) return;
      // Immediate, cheap, synchronous — keeps the cache canonical for reload.
      const stamped = cacheTabsLocally(project.id, next);
      pendingDaemonTabsRef.current = stamped;
      if (tabsDaemonSaveTimerRef.current != null) {
        clearTimeout(tabsDaemonSaveTimerRef.current);
      }
      tabsDaemonSaveTimerRef.current = setTimeout(() => {
        tabsDaemonSaveTimerRef.current = null;
        const pending = pendingDaemonTabsRef.current;
        pendingDaemonTabsRef.current = null;
        if (pending) void persistTabsToDaemonNow(project.id, pending);
      }, TAB_PERSIST_DEBOUNCE_MS);
    },
    [project.id],
  );

  // Flush any pending tab write when the project changes or the view unmounts,
  // so a fast project switch / close doesn't leave the daemon a debounce behind.
  useEffect(() => flushTabsDaemonSave, [flushTabsDaemonSave]);

  const handleActiveWorkspaceContextChange = useCallback((next: WorkspaceContextItem | null) => {
    setActiveWorkspaceContext((current) =>
      workspaceContextItemEqual(current, next) ? current : next,
    );
  }, []);

  const handleWorkspaceContextsChange = useCallback((next: WorkspaceContextItem[]) => {
    setWorkspaceContexts((current) =>
      workspaceContextItemsEqual(current, next) ? current : next,
    );
  }, []);

  const refreshProjectFiles = useCallback(async (): Promise<ProjectFile[]> => {
    const next = await fetchProjectFiles(project.id);
    projectFilesRef.current = next;
    setProjectFiles(next);
    return next;
  }, [project.id]);

  useEffect(() => {
    projectFilesRef.current = projectFiles;
  }, [projectFiles]);

  // Cache HTML file contents so the auto-open module check (issue #2744) does
  // not re-fetch unchanged entries on every Write. Keyed by file name with the
  // mtime stored alongside, so a rewrite REPLACES the file's single entry
  // rather than accreting a new key. Bounded by the project's HTML file count.
  const htmlContentCacheRef = useRef<Map<string, { mtime: number; text: string | null }>>(
    new Map(),
  );
  const readProjectHtml = useCallback(
    async (name: string): Promise<string | null> => {
      const file = projectFilesRef.current.find((entry) => entry.name === name);
      const mtime = file?.mtime ?? 0;
      const cached = htmlContentCacheRef.current.get(name);
      if (cached && cached.mtime === mtime) return cached.text;
      try {
        const response = await fetch(projectRawUrl(project.id, name));
        const text = response.ok ? await response.text() : null;
        htmlContentCacheRef.current.set(name, { mtime, text });
        return text;
      } catch {
        htmlContentCacheRef.current.set(name, { mtime, text: null });
        return null;
      }
    },
    [project.id],
  );

  const refreshLiveArtifacts = useCallback(async (): Promise<LiveArtifactSummary[]> => {
    const next = await fetchLiveArtifacts(project.id);
    setLiveArtifacts(next);
    return next;
  }, [project.id]);

  const refreshWorkspaceItems = useCallback(async (): Promise<ProjectFile[]> => {
    const [nextFiles] = await Promise.all([refreshProjectFiles(), refreshLiveArtifacts()]);
    return nextFiles;
  }, [refreshLiveArtifacts, refreshProjectFiles]);

  useEffect(() => {
    if (!tabsLoadedRef.current) return;
    if (hasAppliedInitialPrimaryOpenRef.current) return;
    if (routeFileName) return;
    if (openTabsState.active || openTabsState.tabs.length > 0) {
      hasAppliedInitialPrimaryOpenRef.current = true;
      return;
    }
    if (tabsHydratedFromSavedStateRef.current) {
      hasAppliedInitialPrimaryOpenRef.current = true;
      return;
    }
    const primaryFile = selectPrimaryProjectFile(projectFiles);
    if (!primaryFile) return;
    hasAppliedInitialPrimaryOpenRef.current = true;
    persistTabsState({ tabs: [primaryFile.name], active: primaryFile.name });
  }, [openTabsState.active, openTabsState.tabs.length, persistTabsState, projectFiles, routeFileName]);

  const requestOpenFile = useCallback((name: string) => {
    if (!name) return;
    setOpenRequest({ name, nonce: Date.now() });
  }, []);

  const persistArtifact = useCallback(
    async (
      art: Artifact,
      projectFilesSnapshot?: ProjectFile[],
      sourceText?: string,
      options: { pointerMinMtime?: number } = {},
    ) => {
      const persistedHtml = resolvePersistedArtifactHtml({
        artifactHtml: art.html,
        identifier: art.identifier,
        sourceText,
      });
      const artifactToPersist = persistedHtml === art.html ? art : { ...art, html: persistedHtml };
      const baseName = artifactBaseNameFor(art);
      const ext = artifactExtensionFor(art);
      // Pick a name that doesn't collide with an existing project file.
      // The first run uses `<base>.<ext>`; subsequent runs append `-2`, `-3`…
      // so prior artifacts aren't silently overwritten.
      const currentProjectFiles = projectFilesSnapshot ?? projectFilesRef.current;
      const existing = new Set(currentProjectFiles.map((f) => f.name));
      let fileName = `${baseName}${ext}`;
      let n = 2;
      while (existing.has(fileName) && savedArtifactRef.current !== fileName) {
        fileName = `${baseName}-${n}${ext}`;
        n += 1;
      }
      if (ext === '.html') {
        const pointerProjectFiles = filterProjectFilesByMinMtime(
          currentProjectFiles,
          options.pointerMinMtime,
        );
        const pointerTarget = resolveHtmlPointerArtifactTarget({
          content: artifactToPersist.html,
          candidateFileName: fileName,
          projectFiles: pointerProjectFiles,
        });
        if (pointerTarget) {
          if (savedArtifactRef.current === pointerTarget) return;
          savedArtifactRef.current = pointerTarget;
          requestOpenFile(pointerTarget);
          return;
        }
      }
      // Pre-write structural gate for HTML artifacts (#50, #1143). Reject
      // bodies that obviously aren't a complete document — usually a one-line
      // prose summary the model emitted inside `<artifact type="text/html">`
      // when only Edit-tool changes happened this turn. Without this guard,
      // such content lands as a phantom HTML file in the project panel.
      if (ext === '.html') {
        const validation = validateHtmlArtifact(artifactToPersist.html);
        if (!validation.ok) {
          setError(`Refused to save artifact "${art.identifier || art.title || 'untitled'}": ${validation.reason}`);
          return;
        }
      }
      if (savedArtifactRef.current === fileName) return;
      const title = art.title || art.identifier || fileName;
      const metadata = {
        identifier: art.identifier,
        artifactType: art.artifactType,
        inferred: false,
      };
      const manifest =
        ext === '.html'
          ? createHtmlArtifactManifest({
              entry: fileName,
              title,
              sourceSkillId: project.skillId ?? undefined,
              designSystemId: project.designSystemId,
              metadata,
            })
          : inferLegacyManifest({
              entry: fileName,
              title,
              metadata: {
                ...metadata,
                sourceSkillId: project.skillId ?? undefined,
                designSystemId: project.designSystemId,
              },
            });
      const file = await writeProjectTextFile(project.id, fileName, artifactToPersist.html, {
        artifactManifest: manifest ?? undefined,
      });
      if (file) {
        savedArtifactRef.current = file.name;
        setFilesRefresh((n) => n + 1);
        // Surface the daemon's stub-guard warning when it fires in `warn`
        // mode (the default). Without this the warning would land in the
        // file metadata silently and the user would never see that the
        // model shipped a placeholder.
        if (file.stubGuardWarning) {
          setError(
            `Saved "${file.name}", but the model may have shipped a placeholder: ` +
              `${file.stubGuardWarning.message}`,
          );
        }
        // Auto-open the freshly-persisted artifact as a tab so the user
        // sees it without an extra click. The Write-tool path already does
        // this for tool-emitted files; this handles the artifact-tag path.
        requestOpenFile(file.name);
      } else {
        // writeProjectTextFile collapses all failure paths (non-OK HTTP
        // responses, network errors, and stub-guard 422s) to null — the
        // helper's return contract would need to be widened to distinguish
        // them, which is out of scope here.  Show a generic banner so the
        // failure is observable rather than silent; the daemon logs carry
        // the structured details for any specific error type.
        // Clear the saved-artifact ref so the user can retry.
        savedArtifactRef.current = '';
        setError(
          `Couldn't save artifact "${fileName}". The write failed — ` +
            'check the daemon logs for details.',
        );
      }
    },
    [project.id, project.designSystemId, project.skillId, requestOpenFile],
  );

  const artifactFromStandaloneHtml = useCallback(
    (sourceText: string): Artifact | null => artifactFromRecoverableSourceText(sourceText),
    [],
  );

  // Set of project file names that the chat surface uses to decide whether
  // a tool card's path is openable as a tab. Recomputed on every file-list
  // change; tool cards just read from the set.
  const projectFileNames = useMemo(
    () => new Set(projectFiles.map((f) => f.name)),
    [projectFiles],
  );
  const activeProjectFileName = useMemo(
    () => (
      openTabsState.active && projectFileNames.has(openTabsState.active)
        ? openTabsState.active
        : null
    ),
    [openTabsState.active, projectFileNames],
  );
  const agentsById = useMemo(
    () => new Map(agents.map((agent) => [agent.id, agent])),
    [agents],
  );

  // Keep the @-picker's source of truth fresh: every refreshSignal bump
  // (artifact saved, sketch saved, image uploaded) refetches; on first
  // mount we also do an initial pull so attachments staged before the
  // agent has written anything still see the user's pasted images.
  useEffect(() => {
    void refreshWorkspaceItems().catch(() => {
      // The daemon probe can briefly lag behind a just-started local
      // runtime. Retry when daemonLive flips or the explicit refresh key
      // changes instead of leaving the project view in its empty shell.
    });
  }, [daemonLive, refreshWorkspaceItems, filesRefresh]);

  // Live-reload: when the daemon's chokidar watcher reports a file change,
  // bump filesRefresh so the file list refetches with new mtimes — which
  // propagates through to FileViewer iframes via PR #384's ?v=${mtime}
  // cache-bust, triggering an automatic preview reload without a click.
  //
  // Coalesce the refresh: agent rewrites surface to chokidar as an
  // `unlink` + `add` (+ later `change`) burst within a single tick (#2195).
  // Refreshing the file list on the intermediate `unlink` makes the open
  // tab's active file vanish for one frame before the `add` restores it,
  // and FileWorkspace's "tab no longer on disk" path then drops the user
  // out of their preview. A short trailing wait absorbs the burst; the
  // maxWait cap stops a sustained edit storm from starving the UI.
  const refreshFilesAndDesignMd = useCallback(() => {
    setFilesRefresh((n) => n + 1);
    // Round 7 (mrcfps): file mutations are the dominant staleness signal
    // post-finalize — bump the refresh key so DESIGN.md staleness
    // recomputes against the new mtimes.
    setDesignMdRefreshKey((n) => n + 1);
  }, []);
  const coalescedFileChangedRefresh = useCoalescedCallback(
    refreshFilesAndDesignMd,
    { wait: 80, maxWait: 250 },
  );
  const handleProjectEvent = useCallback((evt: ProjectEvent) => {
    if (evt.type === 'file-changed') {
      iframeKeepAlivePool.evictProject(project.id);
      coalescedFileChangedRefresh();
      return;
    }
    if (evt.type === 'conversation-created') {
      // A new conversation was inserted into this project by a path the
      // open project view can't observe through its own state (currently:
      // Routines "Run now" in reuse-an-existing-project mode, #1361).
      // Refetch the conversation list so the new entry becomes visible
      // without requiring the user to leave and re-enter the project.
      // Deliberately do NOT change the active conversation here — the
      // user keeps their current context. Auto-switch is a separate UX
      // decision tracked in #1361.
      if (evt.projectId !== project.id) return;
      const capturedProjectId = project.id;
      const myToken = ++conversationsRefreshTokenRef.current;
      void (async () => {
        try {
          const list = await listConversations(capturedProjectId);
          // Bail if the user switched projects while this request was in
          // flight (#1361 review, Codex P1). The captured project id is the
          // one we asked the daemon about; the live ref is the one the
          // user is looking at right now. If they don't match, applying
          // the list would overwrite the new project's sidebar with
          // stale data from the old one.
          if (projectIdRef.current !== capturedProjectId) return;
          // Bail if a newer conversation-created event already dispatched
          // its own refresh after us (#1361 review, lefarcen P2). With two
          // rapid events the later request may resolve first; if this
          // earlier request resolves afterwards it would drop the newer
          // conversation. Only the latest dispatch is allowed to apply.
          if (conversationsRefreshTokenRef.current !== myToken) return;
          setConversations(list);
        } catch {
          // Defensive: refresh failed (network blip, daemon gone). The
          // next project mount or another conversation-created event
          // will retry; no need to surface an error here.
        }
      })();
      return;
    }
    const agentEvent = projectEventToAgentEvent(evt);
    if (!agentEvent) return;
    setLiveArtifactEvents((prev) => appendLiveArtifactEventItem(prev, agentEvent));
    void refreshLiveArtifacts();
    onProjectsRefresh();
    // Live artifact events come from chat-turn-emitted artifacts; they
    // also imply the conversation transcript changed.
    setDesignMdRefreshKey((n) => n + 1);
  }, [coalescedFileChangedRefresh, iframeKeepAlivePool, onProjectsRefresh, refreshLiveArtifacts, project.id]);
  useProjectFileEvents(project.id, daemonLive, handleProjectEvent);

  const activePromptContextSignature = useMemo(() => {
    const skill = project.skillId
      ? (skills.find((s) => s.id === project.skillId) ??
        designTemplates.find((s) => s.id === project.skillId))
      : null;
    const designSystem = project.designSystemId
      ? designSystems.find((d) => d.id === project.designSystemId)
      : null;
    return JSON.stringify({
      designSystem: designSystem
        ? {
            id: designSystem.id,
            title: designSystem.title,
            category: designSystem.category,
            summary: designSystem.summary,
            source: designSystem.source ?? null,
          }
        : null,
      skill: skill
        ? {
            id: skill.id,
            name: skill.name,
            description: skill.description,
            mode: skill.mode,
            source: skill.source ?? null,
            upstream: skill.upstream,
          }
        : null,
    });
  }, [designSystems, designTemplates, project.designSystemId, project.skillId, skills]);
  const previousPromptContextSignatureRef = useRef(activePromptContextSignature);
  useEffect(() => {
    if (previousPromptContextSignatureRef.current === activePromptContextSignature) return;
    previousPromptContextSignatureRef.current = activePromptContextSignature;
    iframeKeepAlivePool.evictProject(project.id, { includeActive: true });
  }, [activePromptContextSignature, iframeKeepAlivePool, project.id]);

  // When the URL points at a specific file, fire an open request so the
  // FileWorkspace promotes it to an active tab. We watch routeFileName
  // (the parsed segment) so back/forward navigation triggers the same path.
  useEffect(() => {
    if (!routeFileName) return;
    requestOpenFile(routeFileName);
  }, [routeFileName, requestOpenFile]);

  // Sync the URL when the active tab changes, so reload + share-link both
  // land back on the same view. Replace (not push) on tab activation so the
  // history stack doesn't fill with every tab click.
  // Composite sync key: tracks BOTH the active file target AND the active
  // conversation id, so a conversation-only change (e.g. `listConversations`
  // resolves after `loadTabs` hydrated the active tab, or the user picks a
  // different conversation under the same tab) still triggers the navigate
  // and pushes `/conversations/:cid` into the URL. Keying only on the file
  // target lost that update because the early-return saw `target` unchanged
  // and skipped the navigate (lefarcen P1 on PR #1508).
  const lastSyncedRouteKeyRef = useRef<string | null>(null);
  const lastSeenRouteConversationIdRef = useRef<string | null>(null);
  useEffect(() => {
    const target = openTabsState.active && (
      openTabsState.tabs.includes(openTabsState.active)
      || projectFileNames.has(openTabsState.active)
      || isLiveArtifactTabId(openTabsState.active)
    )
      ? openTabsState.active
      : null;
    const nextKey = `${activeConversationId ?? ''}:${target ?? ''}`;
    if (nextKey === lastSyncedRouteKeyRef.current) return;
    lastSyncedRouteKeyRef.current = nextKey;
    lastSyncedConversationIdRef.current = activeConversationId;
    // PerishCode + Codex P1 on PR #1508: the prior version of this
    // sync stripped any `/conversations/:cid` segment from the URL as
    // soon as a tab became active, which regressed the deep-link
    // behavior the parent commit was meant to add (reload / share
    // would fall back to `list[0]` instead of the routed run's
    // conversation). Thread the active conversation id so the URL
    // always reflects the conversation the project view is actually
    // showing, matching how `fileName` already tracks the active tab.
    navigate(
      {
        kind: 'project',
        projectId: project.id,
        conversationId: activeConversationId,
        fileName: target,
      },
      { replace: true },
    );
  }, [openTabsState.active, projectFileNames, project.id, activeConversationId]);

  const handleEnsureProject = useCallback(async (): Promise<string | null> => {
    return project.id;
  }, [project.id]);

  const composedSystemPrompt = useCallback(async (
    sessionModeOverride: ChatSessionMode = activeSessionMode,
  ): Promise<string> => {
    let skillBody: string | undefined;
    let skillName: string | undefined;
    let skillMode: SkillSummary['mode'] | undefined;
    let designSystemBody: string | undefined;
    let designSystemTitle: string | undefined;

    if (project.skillId) {
      // project.skillId can resolve to either root after the
      // skills/design-templates split; check both lists so a template-backed
      // project keeps composing its template body when running in API mode.
      const summary =
        skills.find((s) => s.id === project.skillId) ??
        designTemplates.find((s) => s.id === project.skillId);
      skillName = summary?.name;
      skillMode = summary?.mode;
      const cached = skillCache.current.get(project.skillId);
      if (cached !== undefined) {
        skillBody = cached;
      } else {
        const detail =
          (await fetchSkill(project.skillId)) ??
          (await fetchDesignTemplate(project.skillId));
        if (detail) {
          skillBody = detail.body;
          skillCache.current.set(project.skillId, detail.body);
        }
      }
    }
    if (project.designSystemId) {
      const summary = designSystems.find((d) => d.id === project.designSystemId);
      designSystemTitle = summary?.title;
      const cached = designCache.current.get(project.designSystemId);
      if (cached !== undefined) {
        designSystemBody = cached;
      } else {
        const detail = await fetchDesignSystem(project.designSystemId);
        if (detail) {
          designSystemBody = detail.body;
          designCache.current.set(project.designSystemId, detail.body);
        }
      }
    }
    let template: ProjectTemplate | undefined;
    const tplId = project.metadata?.templateId;
    if (project.metadata?.kind === 'template' && tplId) {
      const cached = templateCache.current.get(tplId);
      if (cached) {
        template = cached;
      } else {
        const fetched = await getTemplate(tplId);
        if (fetched) {
          templateCache.current.set(tplId, fetched);
          template = fetched;
        }
      }
    }
    // Fold in the auto-memory block so BYOK / API-mode chats see the
    // same Personal-memory section a daemon-side CLI chat would. The
    // daemon does this by calling `composeMemoryBody()` directly; the
    // web side hits the equivalent HTTP surface so it can stay
    // ignorant of daemon internals. Failures are swallowed — memory is
    // best-effort, never a blocker for the chat round-trip.
    let memoryBody: string | undefined;
    try {
      const resp = await fetch('/api/memory/system-prompt');
      if (resp.ok) {
        const json = (await resp.json()) as MemorySystemPromptResponse;
        if (typeof json.body === 'string' && json.body.trim().length > 0) {
          memoryBody = json.body;
        }
      }
    } catch {
      // Ignore; memory injection is best-effort.
    }
    let audioVoiceOptions: AudioVoiceOption[] | undefined;
    let audioVoiceOptionsLookupError: string | undefined;
    if (shouldFetchElevenLabsVoiceOptions(project)) {
      try {
        audioVoiceOptions = await fetchElevenLabsVoiceOptions();
        setAudioVoiceOptionsError(null);
      } catch (err) {
        const message = err instanceof Error
          ? err.message
          : 'ElevenLabs voice list could not be loaded.';
        audioVoiceOptionsLookupError = message;
        setAudioVoiceOptionsError(message);
      }
    } else {
      setAudioVoiceOptionsError(null);
    }
    return composeSystemPrompt({
      skillBody,
      skillName,
      skillMode,
      designSystemBody,
      designSystemTitle,
      memoryBody,
      metadata: project.metadata,
      template,
      audioVoiceOptions,
      audioVoiceOptionsError: audioVoiceOptionsLookupError,
      streamFormat: config.mode === 'api' ? 'plain' : undefined,
      sessionMode: sessionModeOverride,
      locale,
      userInstructions: config.customInstructions,
    });
  }, [
    project.skillId,
    project.designSystemId,
    project.metadata,
    skills,
    designTemplates,
    designSystems,
    config.mode,
    config.customInstructions,
    activeSessionMode,
    locale,
  ]);

  const persistMessage = useCallback(
    (m: ChatMessage, options?: SaveMessageOptions) => {
      if (!activeConversationId) return;
      // Source-level guard against the "Working 24m+ / Waiting for first
      // output" UI: never write a daemon assistant row that is still
      // queued/running but has no runId. Until POST /api/runs returns the
      // runId, the message is purely in-flight on the client; persisting it
      // here creates a row that nothing can ever reattach to (daemon never
      // saw the runId, client lost the response). Once onRunCreated assigns
      // a runId — or the run finishes terminally — this guard lets the row
      // through normally.
      if (isPhantomDaemonRunMessage(m)) return;
      void saveMessage(project.id, activeConversationId, m, options);
    },
    [project.id, activeConversationId],
  );

  const persistMessageById = useCallback(
    (messageId: string, options?: SaveMessageOptions) => {
      if (!activeConversationId) return;
      setMessages((curr) => {
        const found = curr.find((m) => m.id === messageId);
        if (found && !isPhantomDaemonRunMessage(found)) {
          void saveMessage(project.id, activeConversationId, found, options);
        }
        return curr;
      });
    },
    [project.id, activeConversationId],
  );

  const updateMessageById = useCallback(
    (
      messageId: string,
      updater: (message: ChatMessage) => ChatMessage,
      persist = false,
      persistOptions?: SaveMessageOptions,
    ) => {
      setMessages((curr) => {
        let saved: ChatMessage | null = null;
        const next = curr.map((m) => {
          if (m.id !== messageId) return m;
          const updated = updater(m);
          saved = updated;
          return updated;
        });
        // Same phantom guard as persistMessage: skip writes for a daemon
        // assistant row that is still in-flight (active runStatus, no runId).
        // The runId-arriving update from onRunCreated passes through because
        // the updater sets runId before this check runs.
        if (persist && saved && activeConversationId && !isPhantomDaemonRunMessage(saved)) {
          void saveMessage(project.id, activeConversationId, saved, persistOptions);
        }
        return next;
      });
    },
    [project.id, activeConversationId],
  );

  const appendConversationMessage = useCallback(
    (
      conversationId: string,
      message: ChatMessage,
      options?: SaveMessageOptions,
      persist = true,
    ) => {
      if (
        activeConversationId === conversationId
        || messagesConversationIdRef.current === conversationId
      ) {
        setMessages((curr) => [...curr, message]);
      }
      if (persist) void saveMessage(project.id, conversationId, message, options);
    },
    [activeConversationId, project.id],
  );

  const replaceConversationMessage = useCallback(
    (
      conversationId: string,
      message: ChatMessage,
      options?: SaveMessageOptions,
      persist = true,
    ) => {
      if (
        activeConversationId === conversationId
        || messagesConversationIdRef.current === conversationId
      ) {
        setMessages((curr) => curr.map((item) => (item.id === message.id ? message : item)));
      }
      if (persist) void saveMessage(project.id, conversationId, message, options);
    },
    [activeConversationId, project.id],
  );

  const refreshConversationMessagesFromServer = useCallback(
    async (conversationId: string) => {
      if (messagesConversationIdRef.current !== conversationId) return;
      try {
        const serverMessages = await listMessages(project.id, conversationId);
        if (messagesConversationIdRef.current !== conversationId) return;
        setMessages((current) => mergeServerMessagesIntoConversation(current, serverMessages));
        setMessagesInitialized(true);
        setMessagesConversationId(conversationId);
        setFailedMessagesConversationId(null);
      } catch (err) {
        console.warn('Failed to refresh conversation messages after run completion', err);
      }
    },
    [project.id],
  );

  const scheduleConversationMessageRefresh = useCallback(
    (conversationId: string) => {
      scheduleProjectTimeout(() => {
        void refreshConversationMessagesFromServer(conversationId);
      }, 150);
    },
    [refreshConversationMessagesFromServer, scheduleProjectTimeout],
  );

  const markStreamingConversation = useCallback((conversationId: string) => {
    streamingConversationIdRef.current = conversationId;
    setStreaming(true);
    setStreamingConversationId(conversationId);
  }, []);

  const clearStreamingMarker = useCallback((conversationId?: string | null) => {
    const next = clearStreamingConversationMarker(
      streamingConversationIdRef.current,
      conversationId,
    );
    if (next === streamingConversationIdRef.current) return;
    streamingConversationIdRef.current = next;
    setStreamingConversationId(next);
    setStreaming(next !== null);
  }, []);

  const clearActiveRunRefs = useCallback((
    conversationId: string,
    controller: AbortController,
    cancelController: AbortController,
  ) => {
    if (!shouldClearActiveRunRefs(streamingConversationIdRef.current, conversationId)) {
      return false;
    }
    if (abortRef.current !== controller || cancelRef.current !== cancelController) {
      return false;
    }
    abortRef.current = null;
    cancelRef.current = null;
    return true;
  }, []);

  const clearCurrentRunStreamingMarker = useCallback((
    conversationId: string,
    controller: AbortController,
    cancelController: AbortController,
  ) => {
    if (!clearActiveRunRefs(conversationId, controller, cancelController)) return false;
    clearStreamingMarker(conversationId);
    return true;
  }, [clearActiveRunRefs, clearStreamingMarker]);

  const handleAssistantFeedback = useCallback(
    (assistantMessage: ChatMessage, change: ChatMessageFeedbackChange) => {
      const now = Date.now();
      updateMessageById(
        assistantMessage.id,
        (prev) =>
          change
            ? {
                ...prev,
                feedback: {
                  rating: change.rating,
                  reasonCodes: change.reasonCodes,
                  customReason: change.customReason,
                  reasonsSubmittedAt: change.reasonsSubmittedAt,
                  createdAt:
                    prev.feedback?.rating === change.rating
                      ? prev.feedback.createdAt
                      : now,
                  updatedAt: now,
                },
              }
            : {
                ...prev,
                feedback: undefined,
              },
        true,
      );
      // Forward affirmative ratings to the daemon → Langfuse `score-create`.
      // Clears (change=null) are skipped — Langfuse scores are append-only,
      // and the rating is also captured by the PostHog event so a clear is
      // recoverable downstream if we ever need it.
      const runId = assistantMessage.runId;
      if (change && runId && activeConversationId) {
        void reportChatRunFeedback({
          runId,
          projectId: project.id,
          conversationId: activeConversationId,
          assistantMessageId: assistantMessage.id,
          rating: change.rating,
          reasonCodes: change.reasonCodes ?? [],
          hasCustomReason: !!change.customReason,
          customReason: normalizeCustomReason(change.customReason),
        });
      }
    },
    [updateMessageById, activeConversationId, project.id],
  );

  // `code` is the structured API error code (e.g. AGENT_AUTH_REQUIRED); it
  // rides along on the error status event so AssistantMessage can render the
  // hosted-AMR nudge for model/auth/quota failures on non-AMR agents.
  const appendAssistantErrorEvent = useCallback(
    (messageId: string, message: string, code?: string) => {
      if (!message) return;
      updateMessageById(
        messageId,
        (prev) => appendErrorStatusEvent(prev, message, code),
        true,
      );
    },
    [updateMessageById],
  );

  const auditDesignSystemWorkspaceAfterRun = useCallback(
    async (assistantMessageId: string) => {
      if (!isDesignSystemWorkspaceMetadata(project.metadata)) return;
      try {
        const audit = await fetchProjectDesignSystemPackageAudit(project.id);
        if (!audit) return;
        const auditSummary = summarizeDesignSystemPackageAudit(audit);
        updateMessageById(
          assistantMessageId,
          (prev) => ({
            ...prev,
            events: [...(prev.events ?? []), { kind: 'status', label: 'audit', detail: auditSummary }],
          }),
          true,
          { telemetryFinalized: true },
        );
        const repairPrompt = buildDesignSystemPackageAuditRepairPrompt(audit);
        if (repairPrompt) {
          const seed = { id: `audit-${Date.now()}`, value: repairPrompt };
          setChatSeed(seed);
          if (consumeDesignSystemAuditAutoRepair(project.id)) {
            setAutoAuditRepairSeed(seed);
          }
        } else {
          clearDesignSystemAuditAutoRepair(project.id);
        }
      } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        updateMessageById(
          assistantMessageId,
          (prev) => ({
            ...prev,
            events: [
              ...(prev.events ?? []),
              { kind: 'status', label: 'audit', detail: `Package audit could not run: ${detail}` },
            ],
          }),
          true,
          { telemetryFinalized: true },
        );
      }
    },
    [project.id, project.metadata, updateMessageById],
  );

  const refreshPreviewComments = useCallback(async () => {
    if (!activeConversationId) return;
    const next = await fetchPreviewComments(project.id, activeConversationId);
    setPreviewComments(next);
    setAttachedComments((current) =>
      current
        .map((attached) => next.find((comment) => comment.id === attached.id))
        .filter((comment): comment is PreviewComment => Boolean(comment)),
    );
  }, [project.id, activeConversationId]);

  const savePreviewComment = useCallback(
    async (target: PreviewCommentTarget, note: string, attachAfterSave: boolean, images: File[] = []) => {
      if (!activeConversationId) return null;
      // Upload any attached images first so the saved comment carries durable
      // file paths — this is what lets the comment list / re-opened popover
      // re-display the images instead of losing them on echo.
      let uploadedAttachments: PreviewCommentAttachment[] | undefined;
      if (images.length > 0) {
        const result = await uploadProjectFiles(project.id, images);
        if (result.uploaded.length !== images.length) return null;
        uploadedAttachments = result.uploaded.map((file) => ({ path: file.path, name: file.name }));
      }
      const existing = previewComments.find(
        (comment) => comment.filePath === target.filePath && comment.elementId === target.elementId,
      );
      const attachments = mergePreviewCommentAttachments(existing?.attachments, uploadedAttachments);
      const saved = await upsertPreviewComment(project.id, activeConversationId, {
        target,
        note,
        ...(attachments.length > 0 ? { attachments } : {}),
      });
      if (!saved) return null;
      setPreviewComments((current) => mergeSavedPreviewComment(current, saved));
      setAttachedComments((current) =>
        attachAfterSave ? mergeAttachedComments(current, saved) : current.map((comment) => comment.id === saved.id ? saved : comment),
      );
      return saved;
    },
    [project.id, activeConversationId, previewComments],
  );

  const removePreviewComment = useCallback(
    async (commentId: string) => {
      if (!activeConversationId) return;
      const ok = await deletePreviewComment(project.id, activeConversationId, commentId);
      if (!ok) return;
      setPreviewComments((current) => current.filter((comment) => comment.id !== commentId));
      setAttachedComments((current) => removeAttachedComment(current, commentId));
    },
    [project.id, activeConversationId],
  );

  const attachPreviewComment = useCallback((comment: PreviewComment) => {
    setAttachedComments((current) => mergeAttachedComments(current, comment));
  }, []);

  const detachPreviewComment = useCallback((commentId: string) => {
    setAttachedComments((current) => removeAttachedComment(current, commentId));
  }, []);

  const patchAttachedStatuses = useCallback(
    async (attachments: ChatCommentAttachment[], status: PreviewComment['status']) => {
      if (!activeConversationId || attachments.length === 0) return;
      const persistedAttachments = attachments.filter(
        (attachment) => attachment.source !== 'board-batch',
      );
      if (persistedAttachments.length === 0) return;
      setPreviewComments((current) =>
        current.map((comment) =>
          persistedAttachments.some((attachment) => attachment.id === comment.id)
            ? { ...comment, status }
            : comment,
        ),
      );
      await Promise.all(
        persistedAttachments.map((attachment) =>
          patchPreviewCommentStatus(project.id, activeConversationId, attachment.id, status),
        ),
      );
      void refreshPreviewComments();
    },
    [project.id, activeConversationId, refreshPreviewComments],
  );

  useEffect(() => {
    if (config.mode !== 'daemon' || !daemonLive || !activeConversationId || streaming) return;
    let cancelled = false;
    const reattachConversationId = activeConversationId;

    const attachRecoverableRuns = async () => {
      const missingRunIdMessages = messages.filter((m) => {
        if (m.role !== 'assistant' || m.runId) return false;
        return isActiveRunStatus(m.runStatus);
      });
      const activeRuns = missingRunIdMessages.length > 0
        ? await listActiveChatRuns(project.id, reattachConversationId)
        : [];
      const historicalRuns = missingRunIdMessages.length > 0
        ? (await listProjectRuns()).filter(
            (run) => run.projectId === project.id && run.conversationId === reattachConversationId,
          )
        : [];
      if (cancelled) return;
      const activeByMessage = new Map(
        activeRuns
          .filter((run) => run.assistantMessageId)
          .map((run) => [run.assistantMessageId!, run]),
      );
      const historicalByMessage = new Map(
        historicalRuns
          .filter((run) => run.assistantMessageId)
          .map((run) => [run.assistantMessageId!, run]),
      );

      for (const message of messages) {
        if (cancelled) return;
        if (message.role !== 'assistant') continue;

        const needsFullReplay =
          isActiveRunStatus(message.runStatus) || shouldReplayTerminalRunMessage(message);
        if (!needsFullReplay) continue;
        const fallbackRun = !message.runId
          ? activeByMessage.get(message.id) ?? historicalByMessage.get(message.id) ?? null
          : null;
        const runId = message.runId ?? fallbackRun?.id;
        // Self-heal phantom 'running' rows: when the message has no runId
        // and the daemon has no active run mapped to it, the original send
        // POST was lost (daemon restart mid-flight, the user navigated
        // away before /api/runs returned, or a network blip). Leaving the
        // message as 'running' is what produces the "Waiting for first
        // output — Working 24m+" UI the user reported. Mark it failed so
        // the composer is interactive again and the user can re-send.
        if (!runId) {
          updateMessageById(
            message.id,
            (prev) => ({
              ...prev,
              runStatus: 'failed',
              endedAt: prev.endedAt ?? Date.now(),
            }),
            true,
          );
          continue;
        }
        if (reattachControllersRef.current.has(runId)) continue;
        if (completedReattachRunsRef.current.has(runId)) continue;

        if (fallbackRun && !message.runId) {
          updateMessageById(
            message.id,
            (prev) => ({ ...prev, runId, runStatus: fallbackRun.status }),
            true,
          );
        }

        const status = fallbackRun ?? await fetchChatRunStatus(runId);
        if (cancelled) return;
        if (!status) {
          updateMessageById(
            message.id,
            (prev) => ({ ...prev, runStatus: 'failed', endedAt: prev.endedAt ?? Date.now() }),
            true,
          );
          completedReattachRunsRef.current.add(runId);
          continue;
        }
        updateMessageById(
          message.id,
          (prev) => ({
            ...prev,
            runStatus: status.status,
            ...(status.resumable !== undefined ? { resumable: status.resumable } : {}),
          }),
          true,
        );

        if (shouldReplayTerminalRunMessage(message)) {
          const replayedContent = textContentFromAgentEvents(message.events);
          if (replayedContent.trim().length > 0) {
            const parser = createArtifactParser();
            let parsedArtifact: Artifact | null = null;
            let liveHtml = '';
            for (const ev of [...parser.feed(replayedContent), ...parser.flush()]) {
              if (ev.type === 'artifact:start') {
                liveHtml = '';
                parsedArtifact = {
                  identifier: ev.identifier,
                  artifactType: ev.artifactType,
                  title: ev.title,
                  html: '',
                };
                setArtifact(parsedArtifact);
              } else if (ev.type === 'artifact:chunk') {
                liveHtml += ev.delta;
                parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, liveHtml);
                setArtifact((prev) =>
                  artifactWithHtml(prev, ev.identifier, liveHtml),
                );
              } else if (ev.type === 'artifact:end') {
                parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, ev.fullContent);
                setArtifact((prev) =>
                  prev ? artifactWithHtml(prev, ev.identifier, ev.fullContent) : null,
                );
              }
            }

            updateMessageById(
              message.id,
              (prev) => ({
                ...prev,
                content: replayedContent,
                runStatus: resolveSucceededRunStatus(prev.runStatus),
                endedAt: prev.endedAt ?? Date.now(),
              }),
              true,
              { telemetryFinalized: true },
            );

            let nextFiles = await refreshProjectFiles();
            const beforeFileNames = new Set(
              message.preTurnFileNames ?? nextFiles.map((f) => f.name),
            );
            const artifactToPersist = parsedArtifact?.html
              ? parsedArtifact
              : artifactFromStandaloneHtml(replayedContent);
            let recoveredExistingArtifact: ProjectFile | null = null;
            if (artifactToPersist?.html) {
              const runStartedAt = status.createdAt || message.startedAt || message.createdAt;
              recoveredExistingArtifact = findExistingArtifactProjectFile(
                artifactToPersist,
                nextFiles,
                { minMtime: runStartedAt },
              );
              if (recoveredExistingArtifact) {
                savedArtifactRef.current = recoveredExistingArtifact.name;
                requestOpenFile(recoveredExistingArtifact.name);
              } else {
                savedArtifactRef.current = null;
                await persistArtifact(
                  artifactToPersist,
                  nextFiles,
                  replayedContent,
                  { pointerMinMtime: runStartedAt },
                );
                nextFiles = await refreshProjectFiles();
              }
            }
            const diff = computeProducedFiles(beforeFileNames, nextFiles) ?? [];
            const produced = mergeRecoveredArtifact(diff, recoveredExistingArtifact);
            const producedHtmlToOpen = selectAutoOpenProducedHtml(produced);
            if (producedHtmlToOpen) requestOpenFile(producedHtmlToOpen);
            if (produced.length > 0) {
              updateMessageById(
                message.id,
                (prev) => ({ ...prev, producedFiles: produced }),
                true,
                { telemetryFinalized: true },
              );
            }
            await auditDesignSystemWorkspaceAfterRun(message.id);
            completedReattachRunsRef.current.add(runId);
            onProjectsRefresh();
            continue;
          }
        }

        const controller = new AbortController();
        const cancelController = new AbortController();
        reattachControllersRef.current.set(runId, controller);
        reattachCancelControllersRef.current.set(runId, cancelController);
        if (!isTerminalRunStatus(status.status)) {
          abortRef.current = controller;
          cancelRef.current = cancelController;
          markStreamingConversation(reattachConversationId);
        }
        if (needsFullReplay) {
          updateMessageById(
            message.id,
            (prev) => ({ ...prev, content: '', events: [], producedFiles: undefined }),
          );
        }

        let persistTimer: ReturnType<typeof setTimeout> | null = null;
        const persistSoon = () => {
          if (persistTimer) return;
          persistTimer = scheduleProjectTimeout(() => {
            persistTimer = null;
            persistMessageById(message.id);
          }, 500);
        };
        const persistNow = (options?: SaveMessageOptions) => {
          if (persistTimer) {
            clearProjectTimeout(persistTimer);
            persistTimer = null;
          }
          textBuffer.flush();
          persistMessageById(message.id, options);
        };
        const parser = createArtifactParser();
        let parsedArtifact: Artifact | null = null;
        let liveHtml = '';
        let replayedContent = needsFullReplay ? '' : message.content;
        let replayedEvents: AgentEvent[] = needsFullReplay ? [] : [...(message.events ?? [])];
        const applyContentDelta = (delta: string) => {
          for (const ev of parser.feed(delta)) {
            if (ev.type === 'artifact:start') {
              liveHtml = '';
              parsedArtifact = {
                identifier: ev.identifier,
                artifactType: ev.artifactType,
                title: ev.title,
                html: '',
              };
              setArtifact(parsedArtifact);
            } else if (ev.type === 'artifact:chunk') {
              liveHtml += ev.delta;
              parsedArtifact = parsedArtifact
                ? { ...parsedArtifact, html: liveHtml }
                : {
                    identifier: ev.identifier,
                    title: '',
                    html: liveHtml,
                  };
              setArtifact((prev) =>
                prev
                  ? { ...prev, html: liveHtml }
                  : {
                      identifier: ev.identifier,
                      title: '',
                      html: liveHtml,
                    },
              );
            } else if (ev.type === 'artifact:end') {
              parsedArtifact = parsedArtifact
                ? { ...parsedArtifact, html: ev.fullContent }
                : {
                    identifier: ev.identifier,
                    title: '',
                    html: ev.fullContent,
                  };
              setArtifact((prev) => (prev ? { ...prev, html: ev.fullContent } : null));
            }
          }
        };
        if (!needsFullReplay && message.content) {
          applyContentDelta(message.content);
        }
        const textBuffer = createBufferedTextUpdates({
          updateMessage: (updater) => updateMessageById(message.id, updater),
          persistSoon,
          flushAndPersistNow: () => persistNow({ keepalive: true }),
          onContentDelta: applyContentDelta,
        });
        reattachTextBuffersRef.current.add(textBuffer);
        const unregisterTextBuffer = () => {
          reattachTextBuffersRef.current.delete(textBuffer);
        };

        void reattachDaemonRun({
          runId,
          signal: controller.signal,
          cancelSignal: cancelController.signal,
          initialLastEventId: needsFullReplay ? null : message.lastRunEventId ?? null,
          handlers: {
            onDelta: (delta) => {
              replayedContent += delta;
              textBuffer.appendContent(delta);
            },
            onAgentEvent: (ev) => {
              replayedEvents = [...replayedEvents, ev];
              textBuffer.appendEvent(ev);
            },
            onDone: () => {
              // A reattached run interrupted by a "send now" still receives a
              // late onDone from the daemon. Decide ownership first, then bail
              // BEFORE any current-run side effect (committing buffered text,
              // repainting the artifact preview via setArtifact, re-finalizing
              // the message) — only release this run's bookkeeping. See the
              // streamViaDaemon onDone for the ownership rationale.
              const runMayFinalize =
                !supersededRunsRef.current.has(controller);
              if (runMayFinalize) textBuffer.flush();
              textBuffer.cancel();
              unregisterTextBuffer();
              completedReattachRunsRef.current.add(runId);
              reattachControllersRef.current.delete(runId);
              reattachCancelControllersRef.current.delete(runId);
              clearCurrentRunStreamingMarker(reattachConversationId, controller, cancelController);
              if (!runMayFinalize) return;
              for (const ev of parser.flush()) {
                if (ev.type === 'artifact:end') {
                  parsedArtifact = parsedArtifact
                    ? { ...parsedArtifact, html: ev.fullContent }
                    : {
                        identifier: ev.identifier,
                        title: '',
                        html: ev.fullContent,
                      };
                  setArtifact((prev) => (prev ? { ...prev, html: ev.fullContent } : null));
                }
              }
              updateMessageById(
                message.id,
                (prev) => ({
                  ...prev,
                  content: needsFullReplay ? replayedContent : prev.content,
                  events: needsFullReplay ? replayedEvents : prev.events,
                  runStatus: resolveSucceededRunStatus(prev.runStatus),
                  endedAt: prev.endedAt ?? Date.now(),
                }),
                true,
                { telemetryFinalized: true },
              );
              void (async () => {
                const preTurn = message.preTurnFileNames;
                let nextFiles = await refreshProjectFiles();
                // Use the turn-start snapshot when available so reload
                // recovers files produced before the artifact write too;
                // fall back to the current list for legacy messages.
                const beforeFileNames = new Set(preTurn ?? nextFiles.map((f) => f.name));
                let recoveredExistingArtifact: ProjectFile | null = null;
                const artifactToPersist = parsedArtifact?.html
                  ? parsedArtifact
                  : artifactFromStandaloneHtml(replayedContent);
                if (artifactToPersist?.html) {
                  const producedBeforeFallback = computeProducedFiles(beforeFileNames, nextFiles) ?? [];
                  const runStartedAt = status.createdAt || message.startedAt || message.createdAt;
                  recoveredExistingArtifact = findExistingArtifactProjectFile(
                    artifactToPersist,
                    nextFiles,
                    { minMtime: runStartedAt },
                  ) ?? await findSameTurnHtmlWriteForRecoveredArtifact({
                    artifactHtml: resolvePersistedArtifactHtml({
                      artifactHtml: artifactToPersist.html,
                      identifier: artifactToPersist.identifier,
                      sourceText: replayedContent,
                    }),
                    producedFiles: producedBeforeFallback,
                    readProjectHtml,
                  });
                  if (recoveredExistingArtifact) {
                    savedArtifactRef.current = recoveredExistingArtifact.name;
                    requestOpenFile(recoveredExistingArtifact.name);
                  } else {
                    savedArtifactRef.current = null;
                    await persistArtifact(
                      artifactToPersist,
                      nextFiles,
                      replayedContent,
                      { pointerMinMtime: runStartedAt },
                    );
                    nextFiles = await refreshProjectFiles();
                  }
                }
                const diff = computeProducedFiles(beforeFileNames, nextFiles) ?? [];
                const produced = mergeRecoveredArtifact(diff, recoveredExistingArtifact);
                const producedHtmlToOpen = selectAutoOpenProducedHtml(produced);
                if (producedHtmlToOpen) requestOpenFile(producedHtmlToOpen);
                if (produced.length > 0) {
                  updateMessageById(
                    message.id,
                    (prev) => ({ ...prev, producedFiles: produced }),
                    true,
                    { telemetryFinalized: true },
                  );
                }
                await auditDesignSystemWorkspaceAfterRun(message.id);
              })();
              onProjectsRefresh();
            },
            onError: (err) => {
              const errorCode = (err as Error & { code?: string }).code;
              const resumable = (err as Error & { resumable?: boolean }).resumable === true;
              // A superseded reattached run must not paint a global failure
              // banner or re-finalize its message over the replacement run.
              const runMayFinalize =
                !supersededRunsRef.current.has(controller);
              textBuffer.flush();
              textBuffer.cancel();
              unregisterTextBuffer();
              if (runMayFinalize) {
                setError(err.message);
                appendAssistantErrorEvent(message.id, err.message, errorCode);
                updateMessageById(
                  message.id,
                  (prev) => ({
                    ...prev,
                    runStatus: 'failed',
                    endedAt: prev.endedAt ?? Date.now(),
                    resumable,
                  }),
                  true,
                );
                if (artifactFromRecoverableSourceText(replayedContent)) {
                  void (async () => {
                    if (recoveredArtifactMessagesRef.current.has(message.id)) return;
                    const latestRunStatus = await fetchChatRunStatus(runId).catch(() => null);
                    const artifactToPersist = parsedArtifact?.html
                      ? parsedArtifact
                      : artifactFromStandaloneHtml(replayedContent);
                    if (!artifactToPersist?.html) return;
                    let nextFiles = await refreshProjectFiles();
                    const beforeFileNames = new Set(
                      message.preTurnFileNames ?? nextFiles.map((f) => f.name),
                    );
                    const runStartedAt =
                      latestRunStatus?.createdAt || message.startedAt || message.createdAt;
                    let recoveredExistingArtifact = findExistingArtifactProjectFile(
                      artifactToPersist,
                      nextFiles,
                      { minMtime: runStartedAt },
                    );
                    if (recoveredExistingArtifact) {
                      savedArtifactRef.current = recoveredExistingArtifact.name;
                      requestOpenFile(recoveredExistingArtifact.name);
                    } else {
                      savedArtifactRef.current = null;
                      await persistArtifact(
                        artifactToPersist,
                        nextFiles,
                        replayedContent,
                        { pointerMinMtime: runStartedAt },
                      );
                      nextFiles = await refreshProjectFiles();
                      recoveredExistingArtifact = findExistingArtifactProjectFile(
                        artifactToPersist,
                        nextFiles,
                        { minMtime: runStartedAt },
                      );
                    }
                    const diff = computeProducedFiles(beforeFileNames, nextFiles) ?? [];
                    const produced = mergeRecoveredArtifact(diff, recoveredExistingArtifact);
                    if (produced.length > 0) {
                      recoveredArtifactMessagesRef.current.add(message.id);
                    }
                    const producedHtmlToOpen = selectAutoOpenProducedHtml(produced);
                    if (producedHtmlToOpen) requestOpenFile(producedHtmlToOpen);
                    if (latestRunStatus?.status === 'succeeded') setError(null);
                    updateMessageById(
                      message.id,
                      (prev) => ({
                        ...prev,
                        content: replayedContent,
                        producedFiles: produced.length > 0 ? produced : prev.producedFiles,
                        runStatus:
                          latestRunStatus?.status === 'succeeded'
                            ? resolveSucceededRunStatus(prev.runStatus)
                            : prev.runStatus,
                        endedAt: prev.endedAt ?? Date.now(),
                      }),
                      true,
                      { telemetryFinalized: true },
                    );
                    await auditDesignSystemWorkspaceAfterRun(message.id);
                    onProjectsRefresh();
                  })();
                }
              }
              completedReattachRunsRef.current.add(runId);
              reattachControllersRef.current.delete(runId);
              reattachCancelControllersRef.current.delete(runId);
              clearCurrentRunStreamingMarker(reattachConversationId, controller, cancelController);
              persistNow({ telemetryFinalized: true });
              scheduleConversationMessageRefresh(reattachConversationId);
            },
          },
          onRunStatus: (runStatus) => {
            textBuffer.flush();
            updateMessageById(
              message.id,
              (prev) => ({
                ...prev,
                runStatus,
                endedAt: isTerminalRunStatus(runStatus) ? prev.endedAt ?? Date.now() : prev.endedAt,
              }),
              true,
            );
            if (runStatus === 'canceled') {
              textBuffer.cancel();
              unregisterTextBuffer();
              completedReattachRunsRef.current.add(runId);
              reattachControllersRef.current.delete(runId);
              reattachCancelControllersRef.current.delete(runId);
              clearCurrentRunStreamingMarker(reattachConversationId, controller, cancelController);
              persistNow({ telemetryFinalized: true });
            }
            if (isTerminalRunStatus(runStatus)) {
              scheduleConversationMessageRefresh(reattachConversationId);
            }
          },
          onRunEventId: (lastRunEventId) => {
            textBuffer.flush();
            updateMessageById(message.id, (prev) => ({ ...prev, lastRunEventId }));
            persistSoon();
          },
        })
          .catch((err) => {
            // Skip AbortError (expected on interrupt) and any error from a run
            // that was tagged superseded by a send-now interrupt — it must not
            // surface a global failure over the replacement.
            const runMayFinalize =
              !supersededRunsRef.current.has(controller);
            if ((err as Error).name !== 'AbortError' && runMayFinalize) {
              const msg = err instanceof Error ? err.message : String(err);
              setError(msg);
              appendAssistantErrorEvent(message.id, msg);
              updateMessageById(
                message.id,
                (prev) => ({ ...prev, runStatus: 'failed', endedAt: prev.endedAt ?? Date.now() }),
                true,
                { telemetryFinalized: true },
              );
            }
          })
          .finally(() => {
            textBuffer.flush();
            textBuffer.cancel();
            unregisterTextBuffer();
            if (persistTimer) clearProjectTimeout(persistTimer);
            reattachControllersRef.current.delete(runId);
            reattachCancelControllersRef.current.delete(runId);
            clearActiveRunRefs(reattachConversationId, controller, cancelController);
          });
      }
    };

    void attachRecoverableRuns();
    return () => {
      cancelled = true;
    };
  }, [
    daemonLive,
    config.mode,
    activeConversationId,
    streaming,
    messages,
    project.id,
    updateMessageById,
    persistMessageById,
    auditDesignSystemWorkspaceAfterRun,
    markStreamingConversation,
    clearStreamingMarker,
    clearActiveRunRefs,
    clearCurrentRunStreamingMarker,
    clearProjectTimeout,
    refreshProjectFiles,
    readProjectHtml,
    persistArtifact,
    requestOpenFile,
    onProjectsRefresh,
    scheduleProjectTimeout,
    scheduleConversationMessageRefresh,
  ]);

  useEffect(() => {
    if (config.mode !== 'daemon' || !daemonLive || !activeConversationId) return;
    if (!currentConversationHasRecoverableArtifact) return;
    let cancelled = false;
    let recovering = false;

    const recoverArtifacts = async () => {
      if (recovering) return;
      recovering = true;
      try {
        const serverMessages = await listMessages(project.id, activeConversationId).catch(() => []);
        if (cancelled) return;
        const recoveryMessages = serverMessages.length > 0 ? serverMessages : messagesRef.current;
        for (const message of recoveryMessages) {
          if (cancelled) return;
          if (!hasRecoverableArtifactMessage(message)) continue;
          if (recoveredArtifactMessagesRef.current.has(message.id)) continue;
          const runId = message.runId;
          if (!runId) continue;

          const sourceText = message.content.trim().length > 0
            ? message.content
            : textContentFromAgentEvents(message.events);

          const parser = createArtifactParser();
          let parsedArtifact: Artifact | null = null;
          let liveHtml = '';
          for (const ev of [...parser.feed(sourceText), ...parser.flush()]) {
            if (ev.type === 'artifact:start') {
              liveHtml = '';
              parsedArtifact = {
                identifier: ev.identifier,
                artifactType: ev.artifactType,
                title: ev.title,
                html: '',
              };
              setArtifact(parsedArtifact);
            } else if (ev.type === 'artifact:chunk') {
              liveHtml += ev.delta;
              parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, liveHtml);
              setArtifact((prev) =>
                artifactWithHtml(prev, ev.identifier, liveHtml),
              );
            } else if (ev.type === 'artifact:end') {
              parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, ev.fullContent);
              setArtifact((prev) =>
                prev ? artifactWithHtml(prev, ev.identifier, ev.fullContent) : null,
              );
            }
          }

          const artifactToPersist = parsedArtifact?.html
            ? parsedArtifact
            : artifactFromStandaloneHtml(sourceText);
          if (!artifactToPersist?.html) continue;
          const latestRunStatus = await fetchChatRunStatus(runId).catch(() => null);
          let nextFiles = await refreshProjectFiles();
          if (cancelled) return;
          const beforeFileNames = new Set(
            message.preTurnFileNames ?? nextFiles.map((f) => f.name),
          );
          const runStartedAt =
            latestRunStatus?.createdAt || message.startedAt || message.createdAt;
          let recoveredExistingArtifact = findExistingArtifactProjectFile(
            artifactToPersist,
            nextFiles,
            { minMtime: runStartedAt },
          );
          if (recoveredExistingArtifact) {
            savedArtifactRef.current = recoveredExistingArtifact.name;
            requestOpenFile(recoveredExistingArtifact.name);
          } else {
            savedArtifactRef.current = null;
            await persistArtifact(
              artifactToPersist,
              nextFiles,
              sourceText,
              { pointerMinMtime: runStartedAt },
            );
            nextFiles = await refreshProjectFiles();
            recoveredExistingArtifact = findExistingArtifactProjectFile(
              artifactToPersist,
              nextFiles,
              { minMtime: runStartedAt },
            );
          }
          if (cancelled) return;
          const diff = computeProducedFiles(beforeFileNames, nextFiles) ?? [];
          const produced = mergeRecoveredArtifact(diff, recoveredExistingArtifact);
          if (produced.length === 0) {
            continue;
          }
          recoveredArtifactMessagesRef.current.add(message.id);
          const producedHtmlToOpen = selectAutoOpenProducedHtml(produced);
          if (producedHtmlToOpen) requestOpenFile(producedHtmlToOpen);
          updateMessageById(
            message.id,
            (prev) => ({
              ...prev,
              content: sourceText,
              producedFiles: produced,
              runStatus:
                latestRunStatus?.status === 'succeeded'
                  ? 'succeeded'
                  : prev.runStatus,
              endedAt: prev.endedAt ?? Date.now(),
            }),
            true,
            { telemetryFinalized: true },
          );
          await auditDesignSystemWorkspaceAfterRun(message.id);
          scheduleConversationMessageRefresh(activeConversationId);
          onProjectsRefresh();
        }
      } finally {
        recovering = false;
      }
    };

    void recoverArtifacts();
    const interval = window.setInterval(() => {
      void recoverArtifacts();
    }, 1000);

    return () => {
      cancelled = true;
      window.clearInterval(interval);
    };
  }, [
    daemonLive,
    config.mode,
    activeConversationId,
    project.id,
    currentConversationHasRecoverableArtifact,
    artifactFromStandaloneHtml,
    refreshProjectFiles,
    persistArtifact,
    requestOpenFile,
    updateMessageById,
    auditDesignSystemWorkspaceAfterRun,
    scheduleConversationMessageRefresh,
    onProjectsRefresh,
  ]);

  const commitQueuedChatSends = useCallback((next: QueuedChatSend[]) => {
    queuedChatSendsRef.current = next;
    setQueuedChatSends(next);
    saveQueuedChatSends(project.id, next);
  }, [project.id]);

  const enqueueChatSend = useCallback((item: QueuedChatSend) => {
    const next = [...queuedChatSendsRef.current, item];
    commitQueuedChatSends(next);
  }, [commitQueuedChatSends]);

  const removeQueuedChatSend = useCallback((id: string) => {
    const next = queuedChatSendsRef.current.filter((item) => item.id !== id);
    commitQueuedChatSends(next);
  }, [commitQueuedChatSends]);

  const updateQueuedChatSend = useCallback((id: string, update: QueuedChatSendUpdate) => {
    const next = queuedChatSendsRef.current.map((item) => {
      if (item.id !== id) return item;
      const meta = stripQueueOnlyFromMeta(update.meta);
      const updated: QueuedChatSend = {
        ...item,
        prompt: update.prompt,
        attachments: update.attachments,
        commentAttachments: update.commentAttachments,
      };
      if (meta === undefined) delete updated.meta;
      else updated.meta = meta;
      return updated;
    });
    commitQueuedChatSends(next);
  }, [commitQueuedChatSends]);

  const prioritizeQueuedChatSend = useCallback((id: string) => {
    const item = queuedChatSendsRef.current.find((candidate) => candidate.id === id);
    if (!item) return;
    const next = [item, ...queuedChatSendsRef.current.filter((candidate) => candidate.id !== id)];
    commitQueuedChatSends(next);
  }, [commitQueuedChatSends]);

  const reorderCurrentConversationQueuedChatSends = useCallback((orderedIds: string[]) => {
    if (!activeConversationId || orderedIds.length === 0) return;
    const order = new Map(orderedIds.map((id, index) => [id, index]));
    const current = queuedChatSendsRef.current;
    const originalConversationItems = current.filter(
      (item) => item.conversationId === activeConversationId,
    );
    const sortedConversationItems = [...originalConversationItems].sort((a, b) => {
      const aOrder = order.get(a.id) ?? Number.MAX_SAFE_INTEGER;
      const bOrder = order.get(b.id) ?? Number.MAX_SAFE_INTEGER;
      return aOrder - bOrder;
    });
    if (
      sortedConversationItems.every((item, index) => item.id === originalConversationItems[index]?.id)
    ) {
      return;
    }
    let cursor = 0;
    const next = current.map((item) => {
      if (item.conversationId !== activeConversationId) return item;
      return sortedConversationItems[cursor++] ?? item;
    });
    commitQueuedChatSends(next);
  }, [activeConversationId, commitQueuedChatSends]);

  const queueChatSendForCurrentConversation = useCallback((input: {
    attachments: ChatAttachment[];
    commentAttachments: ChatCommentAttachment[];
    conversationId: string;
    meta?: ProjectChatSendMeta;
    prompt: string;
  }) => {
    const queuedMeta = stripQueueOnlyFromMeta(input.meta);
    enqueueChatSend({
      id: randomUUID(),
      conversationId: input.conversationId,
      prompt: input.prompt,
      attachments: input.attachments,
      commentAttachments: input.commentAttachments,
      ...(queuedMeta === undefined ? {} : { meta: queuedMeta }),
      createdAt: Date.now(),
    });
    if (input.commentAttachments.length > 0) {
      const reservedCommentIds = new Set(
        input.commentAttachments
          .filter((attachment) => attachment.source !== 'board-batch')
          .map((attachment) => attachment.id),
      );
      setAttachedComments((current) =>
        current.filter((comment) => !reservedCommentIds.has(comment.id)),
      );
      if (reservedCommentIds.size > 0) {
        setPreviewComments((current) =>
          current.map((comment) =>
            reservedCommentIds.has(comment.id)
              ? { ...comment, status: 'applying' }
              : comment,
          ),
        );
        void Promise.all(
          Array.from(reservedCommentIds, (commentId) =>
            patchPreviewCommentStatus(project.id, input.conversationId, commentId, 'applying'),
          ),
        ).catch(() => {});
      }
    }
  }, [enqueueChatSend, project.id]);

 const handleSend = useCallback(
    async (
      prompt: string, // 用户输入的提示文本
      attachments: ChatAttachment[], // 聊天附件列表（文件、图片等）
      commentAttachments: ChatCommentAttachment[] = commentsToAttachments(attachedComments), // 将附加评论转换为评论附件格式，默认值为转换后的评论附件
      meta?: ProjectChatSendMeta, // 可选的发送元数据，包含会话模式、重试目标等信息
      baseMessages?: ChatMessage[], // 可选的基础消息列表，用于指定消息历史
    ) => {
      // 记录发送时的参数信息，用于调试
      console.log('handleSend: ', { prompt, attachments, commentAttachments, meta, baseMessages });
      
      // 如果没有活跃的会话ID，终止发送操作
      if (!activeConversationId) return false;
      
      // 如果当前消息所属的会话ID与活跃会话ID不匹配，终止发送操作（防止并发问题）
      if (messagesConversationIdRef.current !== activeConversationId) return false;
      
      // 确定运行会话模式：优先使用meta中指定的模式，否则使用当前活跃的会话模式
      const runSessionMode = meta?.sessionMode ?? activeSessionMode;
      
      // 解析重试目标：如果指定了要重试的助手消息ID，在消息列表中查找对应的消息及其上下文
      const retryTarget = meta?.retryOfAssistantId
        ? resolveRetryTarget(messages, meta.retryOfAssistantId)
        : null;
      
      // 如果指定了重试目标但未找到，终止操作（无效的重试请求）
      if (meta?.retryOfAssistantId && !retryTarget) return false;
      
      // 确定运行上下文：优先使用meta中指定的上下文，否则使用重试目标中用户消息的上下文
      const runContext = meta?.context ?? retryTarget?.userMsg.runContext;
      
      // 确定历史基础消息：如果是重试操作，使用重试目标之前的消息；否则使用传入的基础消息或当前消息列表
      const historyBase = retryTarget ? retryTarget.priorMessages : baseMessages ?? messages;
      
      // 验证输入有效性：非重试操作时，如果提示为空、无附件且无评论附件，终止操作
      if (
        !retryTarget &&
        !prompt.trim() &&
        attachments.length === 0 &&
        commentAttachments.length === 0
      ) return false;
      
      // 合并所有附件：将普通附件和评论中的图片附件合并为统一的附件列表
      const effectiveAttachments = mergeChatAttachments(
        attachments,
        ...commentAttachments.map((attachment) =>
          chatAttachmentsFromPreviewCommentImages(attachment.imageAttachments), // 从评论附件中提取图片并转换为聊天附件格式
        ),
      );
      
      // 处理仅加入队列的情况：非重试操作且指定仅加入队列时，将消息放入待发送队列
      if (!retryTarget && meta?.queueOnly) {
        queueChatSendForCurrentConversation({
          conversationId: activeConversationId, // 当前会话ID
          prompt, // 用户提示文本
          attachments: effectiveAttachments, // 合并后的附件列表
          commentAttachments, // 评论附件列表
          meta: { ...(meta ?? {}), sessionMode: runSessionMode }, // 展开元数据并设置会话模式
        });
        return false; // 返回false表示消息已加入队列但未立即发送
      }
      
      // 如果当前会话正忙（有消息正在处理中），将新消息加入队列等待处理
      if (currentConversationBusy) {
        queueChatSendForCurrentConversation({
          conversationId: activeConversationId, // 当前会话ID
          prompt, // 用户提示文本
          attachments: effectiveAttachments, // 合并后的附件列表
          commentAttachments, // 评论附件列表
          meta: { ...(meta ?? {}), sessionMode: runSessionMode }, // 展开元数据并设置会话模式
        });
        return false; // 返回false表示消息已加入队列等待处理
      }
      
      // 清除聊天种子值，准备新的消息生成
      setChatSeed(null);
      
      // 记录当前正在运行的会话ID
      const runConversationId = activeConversationId;
      
      // 清除之前的错误状态，准备新的消息发送
      setError(null);
      
      // 记录消息开始时间戳，用于计算响应时长
      const startedAt = Date.now();
      
      // 构建用户消息对象：如果是重试则使用原有的用户消息，否则创建新的用户消息
      const userMsg: ChatMessage = retryTarget?.userMsg ?? {
        id: randomUUID(), // 生成唯一消息ID
        role: 'user', // 消息角色为用户
        content: prompt, // 消息内容
        createdAt: startedAt, // 创建时间戳
        sessionMode: runSessionMode, // 会话模式
        ...(meta?.appliedPluginSnapshot
          ? { appliedPluginSnapshot: meta.appliedPluginSnapshot } // 如果有应用插件快照，添加到消息中
          : {}),
        ...(runContext ? { runContext } : {}), // 如果有运行上下文，添加到消息中
        attachments: effectiveAttachments.length > 0 ? effectiveAttachments : undefined, // 有效附件列表（空数组时设为undefined）
        commentAttachments: commentAttachments.length > 0 ? commentAttachments : undefined, // 评论附件列表（空数组时设为undefined）
      };
      
      // 获取最终要使用的评论附件列表
      const runCommentAttachments = userMsg.commentAttachments ?? [];
      
      // 合并用户消息中的附件和评论附件中的图片，形成最终发送的附件列表
      const runAttachments = mergeChatAttachments(
        userMsg.attachments ?? [],
        ...runCommentAttachments.map((attachment) =>
          chatAttachmentsFromPreviewCommentImages(attachment.imageAttachments), // 从评论附件中提取图片并转换格式
        ),
      );
      
      // 确定选择的代理：仅在守护进程模式且指定了代理ID时获取代理对象
      const selectedAgent =
        config.mode === 'daemon' && config.agentId
          ? agentsById.get(config.agentId) // 根据代理ID从映射中获取代理对象
          : null;
      
      // 确定选择的代理模型配置：仅在守护进程模式且指定了代理ID时获取对应的模型配置
      const selectedAgentChoice =
        config.mode === 'daemon' && config.agentId
          ? config.agentModels?.[config.agentId] // 获取指定代理的模型选择配置
          : undefined;
      
      // 获取有效的代理模型选择：根据代理对象和模型配置计算出最终要使用的模型
      const effectiveSelectedAgentChoice = effectiveAgentModelChoice(
        selectedAgent,
        selectedAgentChoice,
      );
      
      // 确定助手消息的代理ID：守护进程模式使用配置的代理ID，否则根据API协议获取
      const assistantAgentId =
        config.mode === 'daemon'
          ? config.agentId ?? undefined // 守护进程模式下使用配置的代理ID
          : apiProtocolAgentId(config.apiProtocol); // 其他模式下根据API协议获取代理ID
      
      // 确定助手消息的显示名称：守护进程模式显示代理模型名称，否则显示API协议模型标签
      const assistantAgentName =
        config.mode === 'daemon'
          ? agentModelDisplayName(
              config.agentId, // 代理ID
              selectedAgent?.name, // 代理名称
              effectiveSelectedAgentChoice?.model, // 选中的模型名称
            )
          : apiProtocolModelLabel(config.apiProtocol, config.model); // API协议模式显示模型标签
      
      // 记录当前项目文件的名称列表，用于跟踪对话前的文件状态
      const preTurnFileNames = projectFiles.map((f) => f.name);
      
      // 生成助手消息的唯一ID
      const assistantId = randomUUID();
      
      // 构建初始助手消息对象，内容为空，状态为运行中
      const assistantMsg: ChatMessage = {
        id: assistantId, // 助手消息ID
        role: 'assistant', // 消息角色为助手
        content: '', // 初始内容为空，后续通过事件流填充
        agentId: assistantAgentId, // 助手代理ID
        agentName: assistantAgentName, // 助手代理显示名称
        events: [], // 事件列表初始为空，用于记录消息生成过程
        createdAt: startedAt, // 创建时间与用户消息相同
        runStatus: config.mode === 'daemon' ? 'running' : undefined, // 守护进程模式下标记为运行中状态
        startedAt, // 开始时间戳
        preTurnFileNames, // 对话前的项目文件名称列表
      };
      
      // 保存最新的助手消息引用，用于后续状态更新
      let latestAssistantMsg: ChatMessage = assistantMsg;
      
      // 定义更新会话最新运行状态的函数，用于跟踪消息处理的进度和结果
      const updateConversationLatestRun = (
        status: NonNullable<ChatMessage['runStatus']>, // 运行状态（非空）：running、completed、error等
        endedAt?: number, // 可选的结束时间戳
      ) => {
        // 更新会话列表，找到正在运行的会话并设置其最新运行状态
        setConversations((curr) =>
          curr.map((conversation) =>
            conversation.id === runConversationId // 匹配当前运行的会话ID
              ? {
                  ...conversation, // 保留会话的其他属性
                  updatedAt: endedAt ?? startedAt, // 更新时间：有结束时间则使用结束时间，否则使用开始时间
                  latestRun: {
                    status, // 设置运行状态
                    startedAt, // 开始时间戳
                    ...(endedAt === undefined
                      ? {} // 如果未提供结束时间，不设置结束相关属性
                      : {
                          endedAt, // 结束时间戳
                          durationMs: Math.max(0, endedAt - startedAt), // 计算运行时长（毫秒），确保非负
                        }),
                  },
                }
              : conversation, // 不匹配的会话保持原样
          ),
        );
      };
      // 将当前助手消息ID添加到活跃完成通知运行集合中，用于跟踪正在进行的运行
      activeCompletionNotificationRunsRef.current.add(assistantId);
      
      // 构建下一个历史消息列表：如果是重试，使用重试目标之前的历史加用户消息；否则使用基础历史加用户消息
      const nextHistory = retryTarget
        ? [...retryTarget.priorMessages, userMsg] // 重试：保留之前的历史，追加新的用户消息
        : [...historyBase, userMsg]; // 非重试：在基础历史上追加用户消息
      
      // 构建可见消息列表：如果是重试，在历史后追加保留的尝试记录和助手消息；否则直接追加助手消息
      const nextVisibleMessages = retryTarget
        ? [...nextHistory, ...retryTarget.preservedAttempts, assistantMsg] // 重试：包含之前失败的尝试记录
        : [...nextHistory, assistantMsg]; // 非重试：仅追加助手消息
      
      // 更新UI中的消息列表为新的可见消息
      setMessages(nextVisibleMessages);
      
      // 标记该会话为流式传输状态，用于UI显示加载动画等
      markStreamingConversation(runConversationId);
      
      // 更新会话的最新运行状态：守护进程模式标记为'running'，其他模式标记为'queued'（排队中）
      updateConversationLatestRun(config.mode === 'daemon' ? 'running' : 'queued');
      
      // 清除当前工件（artifact）状态，准备接收新的生成内容
      setArtifact(null);
      
      // 清除已保存工件的引用
      savedArtifactRef.current = null;
      
      // 触发项目触摸事件，更新项目的最后活动时间
      onTouchProject();
      
      // 如果不是重试操作，将用户消息持久化到存储中
      if (!retryTarget) persistMessage(userMsg);
      
      // 有意不在此处持久化助手消息。在守护进程模式下，助手消息以runStatus='running'且无runId的状态开始，
      // 源码级别的守卫将其视为幻影消息——第一次数据库写入发生在下面的onRunCreated中，
      // 当POST /api/runs返回runId时才进行。在API模式下没有runStatus，
      // 缓冲文本路径将在第一个增量内容到达时立即持久化。
      persistMessage(assistantMsg);
      
      // 如果有评论附件，标记它们为"正在应用"状态，并从待处理评论列表中移除
      if (runCommentAttachments.length > 0) {
        // 异步更新评论附件的状态为'applying'（应用中）
        void patchAttachedStatuses(runCommentAttachments, 'applying');
        
        // 创建已消费评论ID的集合，用于过滤
        const consumedCommentIds = new Set(runCommentAttachments.map((attachment) => attachment.id));
        
        // 从附加评论列表中移除已消费的评论
        setAttachedComments((current) =>
          current.filter((comment) => !consumedCommentIds.has(comment.id)),
        );
      }
      
      // 判断是否为对话的第一个轮次：非重试且历史消息列表为空
      const isFirstTurn = !retryTarget && historyBase.length === 0;
      
      // 生成备用首轮标题：如果是设计系统工作区提示，使用专用标题；否则使用提示摘要或截断的提示文本
      const fallbackFirstTurnTitle = isDesignSystemWorkspacePrompt(prompt)
        ? DESIGN_SYSTEM_WORKSPACE_DISPLAY_TITLE // 设计系统工作区的专用显示标题
        : summarizeProjectNameFromPrompt(prompt) || prompt.slice(0, 60).trim(); // 从提示生成摘要或截取前60个字符
      
      // 从提示中生成备用项目名称
      const fallbackProjectName = summarizeProjectNameFromPrompt(prompt);
      
      // 如果是第一个轮次，从提示中派生一个工作标题，使对话在下拉菜单中可识别，无需等待代理往返
      if (isFirstTurn) {
        const title = fallbackFirstTurnTitle; // 使用备用标题
        if (title) {
          // 更新会话列表中的标题
          setConversations((curr) =>
            curr.map((c) =>
              c.id === runConversationId ? { ...c, title } : c, // 匹配当前会话ID则更新标题
            ),
          );
          // 异步持久化会话标题到后端
          void patchConversation(project.id, runConversationId, { title });
        }
        
        const projectName = fallbackProjectName; // 使用备用项目名称
        // 如果项目名称存在、与当前项目名称不同，且允许从提示自动重命名项目
        if (
          projectName &&
          projectName !== project.name &&
          canAutoRenameProjectFromPrompt(project, prompt)
        ) {
          // 构建项目元数据，标记名称来源为'prompt'
          const metadata = project.metadata
            ? { ...project.metadata, nameSource: 'prompt' as const } // 保留原有元数据，添加名称来源标记
            : undefined;
          
          // 创建更新后的项目对象
          const updated: Project = {
            ...project,
            name: projectName, // 更新项目名称
            ...(metadata ? { metadata } : {}), // 如果有元数据则添加
            updatedAt: Date.now(), // 更新时间为当前时间戳
          };
          
          // 通知项目变更
          onProjectChange(updated);
          
          // 异步持久化项目名称到后端
          void patchProject(project.id, {
            name: projectName,
            ...(metadata ? { metadata } : {}),
          });
        }
      }
      
      // 定义判断是否可替换对话标题的函数
      const canReplaceConversationTitle = (title: string | null | undefined) => {
        const trimmed = (title ?? '').trim(); // 去除标题首尾空格
        return (
          !trimmed || // 标题为空
          trimmed === fallbackFirstTurnTitle || // 标题等于备用首轮标题
          trimmed === prompt.slice(0, 60).trim() // 标题等于提示的前60个字符
        );
      };
      
      // 定义应用代理生成标题的函数
      const applyAgentGeneratedTitle = (rawTitle: string) => {
        if (!isFirstTurn) return; // 仅处理首轮对话
        
        const agentTitle = rawTitle.trim(); // 去除代理生成标题的首尾空格
        if (!agentTitle || isDesignSystemWorkspacePrompt(prompt)) return; // 标题为空或是设计系统工作区提示则忽略
        
        // 获取当前对话的标题
        const currentConversationTitle = conversationsRef.current.find(
          (conversation) => conversation.id === runConversationId,
        )?.title;
        
        // 判断是否需要持久化更新对话标题
        const shouldPatchConversation = canReplaceConversationTitle(currentConversationTitle);
        
        // 更新UI中的对话标题
        setConversations((curr) =>
          curr.map((conversation) => {
            if (conversation.id !== runConversationId) return conversation; // 不匹配的会话保持不变
            if (!canReplaceConversationTitle(conversation.title)) return conversation; // 不可替换的标题保持不变
            return { ...conversation, title: agentTitle }; // 更新为代理生成的标题
          }),
        );
        
        // 如果需要持久化，异步更新后端对话标题
        if (shouldPatchConversation) {
          void patchConversation(project.id, runConversationId, { title: agentTitle });
        }
        
        // 如果代理标题与项目名称不同，且允许从提示自动重命名项目
        if (
          agentTitle !== project.name &&
          canAutoRenameProjectFromPrompt(project, prompt)
        ) {
          // 构建项目元数据，标记名称来源为'agent'
          const metadata = project.metadata
            ? { ...project.metadata, nameSource: 'agent' as const } // 保留原有元数据，添加名称来源标记
            : undefined;
          
          // 创建更新后的项目对象
          const updated: Project = {
            ...project,
            name: agentTitle, // 更新项目名称为代理生成的标题
            ...(metadata ? { metadata } : {}), // 如果有元数据则添加
            updatedAt: Date.now(), // 更新时间为当前时间戳
          };
          
          // 通知项目变更
          onProjectChange(updated);
          
          // 异步持久化项目名称到后端
          void patchProject(project.id, {
            name: agentTitle,
            ...(metadata ? { metadata } : {}),
          });
        }
      };

      // 在轮次开始时拍摄文件列表快照，以便在代理完成后进行差异比较，
      // 将新生成的文件（如.pptx）作为助手消息上的下载芯片展示
      const beforeFileNames = new Set(preTurnFileNames);

      // 创建工件解析器，用于从流式文本中提取HTML工件
      const parser = createArtifactParser();
      
      // 初始化解析后的工件对象
      let parsedArtifact: Artifact | null = null;
      
      // 初始化实时HTML内容缓冲区
      let liveHtml = '';
      
      // 初始化流式文本累积器
      let streamedText = '';

      // 定义更新助手消息的辅助函数，接受一个更新函数并应用到消息列表中的助手消息
      const updateAssistant = (updater: (prev: ChatMessage) => ChatMessage) => {
        setMessages((curr) =>
          curr.map((m) => {
            if (m.id !== assistantId) return m; // 非助手消息保持不变
            const updated = updater(m); // 应用更新函数
            latestAssistantMsg = updated; // 更新最新的助手消息引用
            return updated; // 返回更新后的消息
          }),
        );
      };
      
      // 持久化定时器，用于延迟保存助手消息
      let persistTimer: ReturnType<typeof setTimeout> | null = null;
      
      // 定义延迟持久化助手消息的函数，使用500ms防抖
      const persistAssistantSoon = () => {
        if (persistTimer) return; // 如果已有定时器在运行，直接返回
        persistTimer = scheduleProjectTimeout(() => {
          persistTimer = null; // 清除定时器引用
          persistMessageById(assistantId); // 按ID持久化助手消息
        }, 500);
      };
      
      // 定义立即持久化助手消息的函数（保持连接活跃）
      const persistAssistantNowKeepalive = () => {
        if (persistTimer) {
          clearProjectTimeout(persistTimer); // 清除延迟定时器
          persistTimer = null; // 清除定时器引用
        }
        persistMessageById(assistantId, { keepalive: true }); // 立即持久化并保持连接
      };
      
      // 定义推送事件的函数，将代理事件添加到助手消息的事件列表中
      const pushEvent = (ev: AgentEvent) => {
        textBuffer.flush(); // 刷新文本缓冲区，确保事件顺序正确
        updateAssistant((prev) => ({ ...prev, events: [...(prev.events ?? []), ev] })); // 追加事件到消息
        
        // 处理实时工件事件：当收到live_artifact事件时
        if (ev.kind === 'live_artifact') {
          setLiveArtifactEvents((prev) => appendLiveArtifactEventItem(prev, ev)); // 追加实时工件事件
          void refreshLiveArtifacts().then(() => {
            if (ev.action !== 'deleted') requestOpenFile(liveArtifactTabId(ev.artifactId)); // 非删除操作则打开工件文件
          });
          onProjectsRefresh(); // 刷新项目列表
          return;
        }
        
        // 处理实时工件刷新事件
        if (ev.kind === 'live_artifact_refresh') {
          setLiveArtifactEvents((prev) => appendLiveArtifactEventItem(prev, ev)); // 追加刷新事件
          void refreshLiveArtifacts(); // 异步刷新实时工件
          onProjectsRefresh(); // 刷新项目列表
          return;
        }
        
        persistAssistantSoon(); // 延迟持久化助手消息
        persistAssistantSoon(); // 第二次调用确保定时器被设置（可能是冗余的防护）
        
        // 跟踪Write工具调用，以便在代理完成写入文件时自动打开目标文件。
        // 我们关心的文件创建工具：Write（新文件）、Edit（现有文件——展示刚修改的文件也很有用）
        if (ev.kind === 'tool_use') {
          // 权威输入已到达；丢弃实时部分输入，使卡片从解析后的tool_use.input渲染，
          // 而不是从令牌中间的JSON片段渲染
          setLiveToolInput((prev) => {
            if (!(ev.id in prev)) return prev; // 如果ID不在实时输入中，直接返回
            const next = { ...prev };
            delete next[ev.id]; // 删除对应的实时输入项
            return next;
          });
        }
        
        // 处理Write和Edit工具调用，记录文件路径以便后续自动打开
        if (ev.kind === 'tool_use' && ((ev.name === 'Write' || ev.name === 'write') || ev.name === 'Edit')) {
          const input = ev.input as { file_path?: unknown; filePath?: unknown } | null;
          const filePath = input?.file_path ?? input?.filePath; // 获取文件路径（兼容不同字段名）
          if (typeof filePath === 'string' && filePath.length > 0) {
            // 保留完整路径，以便decideAutoOpenAfterWrite可以针对项目的相对文件路径进行路径后缀匹配。
            // 在此处缩减为基本名称会丢失我们在项目树内及外部消除同名冲突所需的段对齐信息
            pendingWritesRef.current.set(ev.id, filePath); // 记录工具调用ID与文件路径的映射
          }
        }
        
        // 处理工具结果事件：当工具执行完成时
        if (ev.kind === 'tool_result') {
          const filePath = pendingWritesRef.current.get(ev.toolUseId); // 获取对应的文件路径
          if (filePath) {
            pendingWritesRef.current.delete(ev.toolUseId); // 删除已处理的映射
            if (!ev.isError) { // 仅处理成功的工具调用
              // 先刷新文件列表，以便FileWorkspace的文件列表（和标签体）在请求聚焦前看到新内容。
              // 仅当文件确实落在项目的文件列表中时才自动打开——
              // 否则项目外的Write（例如上游仓库编辑）将生成一个永久占位标签
              void refreshProjectFiles().then(async (nextFiles) => {
                // .jsx/.tsx文件由同级HTML入口加载时是多文件React原型的模块，不是独立页面——
                // 不要让用户停留在无用的预览标签上。Issue #2744
                const moduleFileNames = /\.(jsx|tsx)$/i.test(filePath)
                  ? await collectReferencedJsxNames(nextFiles, readProjectHtml) // 收集引用的JSX名称
                  : undefined;
                
                // 决定是否在写入后自动打开文件
                const decision = decideAutoOpenAfterWrite(filePath, nextFiles, {
                  moduleFileNames,
                });
                if (decision.shouldOpen && decision.fileName) {
                  requestOpenFile(decision.fileName); // 请求打开文件
                }
              });
            }
          }
        }
      };

      // 定义应用内容增量的函数，处理流式文本中的工件标记
      const applyContentDelta = (delta: string) => {
        for (const ev of parser.feed(delta)) { // 遍历解析器从增量中提取的事件
          if (ev.type === 'artifact:start') { // 工件开始标记
            liveHtml = ''; // 重置实时HTML缓冲区
            parsedArtifact = {
              identifier: ev.identifier, // 工件标识符
              artifactType: ev.artifactType, // 工件类型
              title: ev.title, // 工件标题
              html: '', // 初始HTML为空
            };
            setArtifact(parsedArtifact); // 设置工件状态
          } else if (ev.type === 'artifact:chunk') { // 工件内容块
            liveHtml += ev.delta; // 追加HTML增量到缓冲区
            parsedArtifact = parsedArtifact
              ? { ...parsedArtifact, html: liveHtml } // 更新现有工件的HTML
              : {
                  identifier: ev.identifier,
                  title: '',
                  html: liveHtml,
                }; // 创建新工件对象
            setArtifact((prev) =>
              prev
                ? { ...prev, html: liveHtml } // 更新现有工件
                : {
                    identifier: ev.identifier,
                    title: '',
                    html: liveHtml,
                  }, // 创建新工件
            );
          } else if (ev.type === 'artifact:end') { // 工件结束标记
            parsedArtifact = parsedArtifact
              ? { ...parsedArtifact, html: ev.fullContent } // 使用完整内容更新工件
              : {
                  identifier: ev.identifier,
                  title: '',
                  html: ev.fullContent,
                }; // 创建包含完整内容的工件
            setArtifact((prev) => (prev ? { ...prev, html: ev.fullContent } : null)); // 更新或清除工件状态
          }
        }
      };

      // 创建缓冲文本更新器，用于管理流式文本的缓冲和持久化
      const textBuffer = createBufferedTextUpdates({
        updateMessage: updateAssistant, // 更新助手消息的回调
        persistSoon: persistAssistantSoon, // 延迟持久化回调
        flushAndPersistNow: persistAssistantNowKeepalive, // 立即刷新并持久化回调
        onContentDelta: applyContentDelta, // 内容增量处理回调
      });
      
      // 将文本缓冲器引用保存到ref中，供其他地方使用
      sendTextBufferRef.current = textBuffer;

      // 创建中止控制器，用于取消正在进行的请求
      const controller = new AbortController();
      
      // 创建取消控制器，用于处理用户取消操作
      const cancelController = new AbortController();
      
      // 保存中止和取消控制器的引用
      abortRef.current = controller;
      cancelRef.current = cancelController;
      
      // 定义事件处理器集合
      const handlers = {
        // 处理文本增量：当收到新的文本片段时
        onDelta: (delta: string) => {
          streamedText += delta; // 累积流式文本
          textBuffer.appendContent(delta); // 将增量追加到文本缓冲区
        },
        
        // 处理代理事件：当收到代理生成的事件时
        onAgentEvent: (ev: AgentEvent) => {
          if (ev.kind === 'conversation_title') { // 对话标题事件
            applyAgentGeneratedTitle(ev.title); // 应用代理生成的标题
            return;
          }
          if (ev.kind === 'text') textBuffer.appendTextEvent(ev.text); // 文本事件追加到缓冲区
          else pushEvent(ev); // 其他事件推送到事件列表
        },
        
        // 处理工具输入增量：当工具调用的参数以流式方式到达时
        onToolInputDelta: (id: string, name: string, delta: string) => {
          setLiveToolInput((prev) => ({
            ...prev,
            [id]: {
              name, // 工具名称
              text: (prev[id]?.text ?? '') + delta, // 累积工具输入文本
              // 在首次见到工具时固定其流位置：消息上已有的事件计数是模型在工具调用前发出的所有内容（前导）。
              // 缓冲文本（appendTextEvent）直到下一帧才会刷新到events中，
              // 因此为任何仍在等待的前导块加1——它将在工具位置之前作为一个文本事件提交
              seq:
                prev[id]?.seq ??
                ((latestAssistantMsg.events?.length ?? 0) + (textBuffer.hasPendingText() ? 1 : 0)),
            },
          }));
        },
        
        // 处理完成事件：当代理完成消息生成时
        onDone: (fullText = '') => {
          // 守护进程即使对于已取消的运行也会传递onDone，因此被"立即发送"中断取代的运行
          // 仍可能到达此处，且绝不能将其完成副作用应用于替代运行之上。
          // 运行可以最终化，除非它在中断时被标记为已取代（在handleStop清除引用之前记录），
          // 这即使在替代发送附加之前也是可靠的——不像abortRef，
          // 其最终的onRunStatus/handleStop搅动使其在此处模糊不清
          const runMayFinalize =
            !supersededRunsRef.current.has(controller); // 检查运行是否未被取代
          if (!runMayFinalize) {
            textBuffer.cancel(); // 取消文本缓冲区
            cancelSendTextBuffer(); // 取消发送文本缓冲区
            return;
          }
          
          textBuffer.flush(); // 刷新文本缓冲区
          textBuffer.cancel(); // 取消文本缓冲区
          cancelSendTextBuffer(); // 取消发送文本缓冲区
          
          // 刷新解析器中剩余的内容
          for (const ev of parser.flush()) {
            if (ev.type === 'artifact:end') { // 处理剩余的工件结束事件
              parsedArtifact = parsedArtifact
                ? { ...parsedArtifact, html: ev.fullContent }
                : {
                    identifier: ev.identifier,
                    title: '',
                    html: ev.fullContent,
                  };
              setArtifact((prev) => (prev ? { ...prev, html: ev.fullContent } : null));
            }
          }
          
          // 检查是否为空的API响应：API模式下，无完整文本、无流式文本且无实时HTML
          const emptyApiResponse =
            config.mode === 'api' &&
            !fullText.trim() &&
            !streamedText.trim() &&
            !liveHtml.trim();
          
          if (emptyApiResponse) { // 处理空响应情况
            const endedAt = Date.now();
            const diagnostic = t('assistant.emptyResponseMessage'); // 获取空响应诊断消息
            updateMessageById(
              assistantId,
              (prev) => ({
                ...prev,
                endedAt, // 设置结束时间
                runStatus: 'failed', // 标记为失败状态
                events: [
                  ...(prev.events ?? []),
                  { kind: 'status', label: 'empty_response', detail: config.model }, // 添加空响应状态事件
                  { kind: 'text', text: diagnostic }, // 添加诊断文本事件
                ],
              }),
              true,
              { telemetryFinalized: true }, // 标记遥测已最终化
            );
            
            if (runCommentAttachments.length > 0) {
              void patchAttachedStatuses(runCommentAttachments, 'failed'); // 标记评论附件为失败
            }
            
            // 清除当前运行的流标记，检查是否拥有当前运行
            const ownsCurrentRun = clearCurrentRunStreamingMarker(
              runConversationId,
              controller,
              cancelController,
            );
            if (ownsCurrentRun) updateConversationLatestRun('failed', endedAt); // 更新会话运行状态为失败
            
            void refreshProjectFiles(); // 刷新项目文件
            onProjectsRefresh(); // 刷新项目
            return;
          }
          
          const endedAt = Date.now(); // 记录结束时间
          let finalRunStatus: ChatMessage['runStatus'] = 'succeeded'; // 默认运行状态为成功
          
          // 更新助手消息，设置结束时间和运行状态
          updateAssistant((prev) => {
            finalRunStatus = resolveSucceededRunStatus(prev.runStatus); // 解析最终运行状态
            return {
              ...prev,
              endedAt, // 设置结束时间
              runStatus: finalRunStatus, // 设置运行状态
            };
          });
          
          if (runCommentAttachments.length > 0) {
            void patchAttachedStatuses(runCommentAttachments, 'needs_review'); // 标记评论附件需要审查
          }
          
          // 清除当前运行的流标记
          const ownsCurrentRun = clearCurrentRunStreamingMarker(
            runConversationId,
            controller,
            cancelController,
          );
          if (ownsCurrentRun) updateConversationLatestRun(finalRunStatus ?? 'succeeded', endedAt); // 更新会话运行状态
          
          // 直接重新获取文件列表（而不仅仅是触发刷新信号），以便与轮次前的快照进行差异比较，
          // 并将新文件作为下载芯片附加到助手消息上
          void (async () => {
            let nextFiles = await refreshProjectFiles(); // 刷新项目文件列表
            const finalText = streamedText || fullText; // 获取最终文本内容
            const artifactToPersist = parsedArtifact?.html
              ? parsedArtifact
              : artifactFromStandaloneHtml(finalText); // 确定要持久化的工件
            
            if (artifactToPersist?.html) { // 如果有工件HTML需要持久化
              const producedBeforeFallback = computeProducedFiles(beforeFileNames, nextFiles) ?? []; // 计算已生成的文件
              const sameTurnHtmlWrite = await findSameTurnHtmlWriteForRecoveredArtifact({
                artifactHtml: resolvePersistedArtifactHtml({
                  artifactHtml: artifactToPersist.html,
                  identifier: artifactToPersist.identifier,
                  sourceText: finalText,
                }),
                producedFiles: producedBeforeFallback,
                readProjectHtml,
              });
              
              if (sameTurnHtmlWrite) { // 如果在同一轮次有HTML写入
                savedArtifactRef.current = sameTurnHtmlWrite.name; // 保存工件引用
                requestOpenFile(sameTurnHtmlWrite.name); // 请求打开文件
              } else {
                await persistArtifact(artifactToPersist, nextFiles, finalText); // 持久化工件
                nextFiles = await refreshProjectFiles(); // 重新刷新文件列表
              }
            }
            
            const produced = computeProducedFiles(beforeFileNames, nextFiles) ?? []; // 计算生成的文件
            const producedHtmlToOpen = selectAutoOpenProducedHtml(produced); // 选择要自动打开的HTML文件
            if (producedHtmlToOpen) requestOpenFile(producedHtmlToOpen); // 请求打开HTML文件
            
            // 更新消息列表，附加生成的文件信息
            setMessages((curr) => {
              const updated = curr.map((m) =>
                m.id === assistantId
                  ? { ...m, producedFiles: produced } // 为助手消息添加生成的文件列表
                  : m,
              );
              const finalized = updated.find((m) => m.id === assistantId); // 查找最终的助手消息
              if (finalized) persistMessage(finalized, { telemetryFinalized: true }); // 持久化最终消息并标记遥测完成
              return updated;
            });
            
            await auditDesignSystemWorkspaceAfterRun(assistantId); // 审计设计系统工作区
          })();
          
          onProjectsRefresh(); // 刷新项目
        },
        
        // 处理错误事件：当消息生成过程中发生错误时
        onError: (err: Error) => {
          const endedAt = Date.now(); // 记录错误发生时间
          const errorCode = (err as Error & { code?: string }).code; // 获取错误代码
          const resumable = (err as Error & { resumable?: boolean }).resumable === true; // 检查是否可恢复
          
          // 被"立即发送"中断取代的运行仍可能引发晚期断开错误（例如已取消的流失去了其终端SSE）。
          // 一旦被标记为已取代，它绝不能绘制全局失败横幅或重新最终化其已取消的助手消息。
          // 有关所有权原理，请参见上面的onDone
          const runMayFinalize =
            !supersededRunsRef.current.has(controller); // 检查运行是否未被取代
          
          textBuffer.flush(); // 刷新文本缓冲区
          textBuffer.cancel(); // 取消文本缓冲区
          cancelSendTextBuffer(); // 取消发送文本缓冲区
          
          if (runMayFinalize) { // 如果运行可以最终化
            setError(err.message); // 设置错误消息
            appendAssistantErrorEvent(assistantId, err.message, errorCode); // 追加错误事件到助手消息
            updateAssistant((prev) => ({
              ...prev,
              endedAt, // 设置结束时间
              runStatus: config.mode === 'api' || prev.runId || isActiveRunStatus(prev.runStatus)
                ? 'failed' // API模式、有运行ID或活跃状态时标记为失败
                : prev.runStatus, // 否则保持原状态
              resumable, // 设置可恢复标志
            }));
            
            if (runCommentAttachments.length > 0) {
              void patchAttachedStatuses(runCommentAttachments, 'failed'); // 标记评论附件为失败
            }
          }
          
          // 清除当前运行的流标记
          const ownsCurrentRun = clearCurrentRunStreamingMarker(
            runConversationId,
            controller,
            cancelController,
          );
          if (ownsCurrentRun) updateConversationLatestRun('failed', endedAt); // 更新会话运行状态为失败
          
          // 更新消息列表并持久化
          setMessages((curr) => {
            const finalized = curr.find((m) => m.id === assistantId); // 查找最终的助手消息
            if (finalized) persistMessage(finalized, { telemetryFinalized: true }); // 持久化最终消息
            return curr;
          });
          
          void refreshProjectFiles(); // 刷新项目文件
        },
      };
      // 守护进程模式：通过本地代理处理消息
      if (config.mode === 'daemon') {
        // 如果没有选择代理，抛出错误并终止
        if (!config.agentId) {
          handlers.onError(new Error('Pick a local agent first (top bar).')); // 提示用户先在顶部栏选择本地代理
          return true; // 返回true表示已处理（错误状态）
        }
        
        // 获取有效的代理模型选择配置
        const choice = effectiveSelectedAgentChoice;
        
        // v2版本分析：当活跃项目是设计系统工作区时（由prepareCreatedDesignSystemProject创建，
        // 通过metadata.importedFrom === 'design-system'标识），
        // 从此编辑器启动的每次运行都是设计系统变体运行。
        // 传递analyticsHints以便守护进程在page_name=design_system_project、
        // area=design_system_generation、project_kind=design_system下发出run_created/run_finished事件。
        // 进入设计系统工作区的第一条消息是自动发送的生成启动消息
        // （entry_from='onboarding_design_system'是文档中用于"设计系统创建流程移交给代理"的名称）；
        // 后续消息是审查驱动的重新生成（'regenerate_from_review'）。
        // 使用messages.length === 0判断——比autoSendFirstMessageRef更可靠，
        // 后者会与StrictMode重新挂载和sessionStorage清除产生竞态条件
        const isDesignSystemWorkspaceProject =
          project.metadata?.importedFrom === 'design-system'; // 检查项目是否为设计系统工作区
        
        // 确定设计系统入口来源：消息列表为空则为引导流程，否则为审查重新生成
        const dsEntryFrom: 'onboarding_design_system' | 'regenerate_from_review' =
          messages.length === 0
            ? 'onboarding_design_system' // 首次生成，来源为引导设计系统
            : 'regenerate_from_review'; // 后续生成，来源为审查重新生成
        
        // 构建设计系统分析提示：如果是设计系统工作区项目，则包含相关元数据
        const dsAnalyticsHints = isDesignSystemWorkspaceProject
          ? {
              entryFrom: dsEntryFrom, // 入口来源
              projectKind: 'design_system' as const, // 项目类型为设计系统
              designSystemRunContext: {
                origin: 'manual_create' as const, // 运行上下文：手动创建
              },
            }
          : undefined; // 非设计系统项目则不设置分析提示
        
        // 调用方提供的entry_from（例如来自可恢复失败的'resume_continue'继续操作）
        // 会覆盖设计系统默认值，以便将运行归因于启动它的操作
        //
        // 会话维度提示会在每次实际运行创建时标记（此路径仅对非排队发送运行）：
        // 为此浏览器会话声明下一个从0开始的轮次索引，
        // 并标记项目是否已有生成的工件（项目范围），以便运行被解读为编辑而非首次创建
        const sessionTurn = claimRunTurnIndex(); // 声明并获取当前会话的运行轮次索引
        
        // 检查项目是否已存在工件清单，用于判断是首次创建还是编辑
        const hasExistingArtifact = projectFilesRef.current.some(
          (file) => Boolean(file.artifactManifest), // 检查文件是否有工件清单
        );
        
        // 构建运行分析提示对象，合并设计系统提示、入口来源、会话轮次等信息
        const runAnalyticsHints = {
          ...(dsAnalyticsHints ?? {}), // 合并设计系统分析提示（如果存在）
          ...(meta?.entryFrom ? { entryFrom: meta.entryFrom } : {}), // 如果调用方提供了入口来源，则覆盖默认值
          ...(sessionTurn
            ? { turnIndex: sessionTurn.turnIndex, isFirstRun: sessionTurn.isFirstRun } // 添加轮次索引和是否首次运行的标记
            : {}),
          hasExistingArtifact,
          // This branch only runs in daemon (local-execution) mode, so the
          // runtime is the bundled AMR cloud agent or a local coding CLI —
          // never BYOK (that path streams client-side, below). Hand the daemon
          // the authoritative value so run_created/run_finished split AMR vs
          // CLI without relying on its agent-id re-derivation.
          runtimeType: config.agentId === 'amr' ? ('amr_cloud' as const) : ('local_cli' as const),
        };
        
        // 通过守护进程流式传输消息
        void streamViaDaemon({
          agentId: config.agentId, // 代理ID
          history: nextHistory, // 对话历史
          signal: controller.signal, // 中止信号
          cancelSignal: cancelController.signal, // 取消信号
          handlers, // 事件处理器
          projectId: project.id, // 项目ID
          conversationId: runConversationId, // 会话ID
          assistantMessageId: assistantId, // 助手消息ID
          clientRequestId: randomUUID(), // 客户端请求ID（随机生成）
          skillId: project.skillId ?? null, // 技能ID（如果存在）
          skillIds: Array.isArray(meta?.skillIds) ? meta.skillIds : [], // 技能ID列表（确保为数组）
          context: runContext, // 运行上下文
          designSystemId: project.designSystemId ?? null, // 设计系统ID（如果存在）
          attachments: runAttachments.map((a) => a.path), // 附件路径列表
          commentAttachments: runCommentAttachments, // 评论附件列表
          sessionMode: runSessionMode, // 会话模式
          appliedPluginSnapshotId:
            meta?.appliedPluginSnapshotId ?? meta?.appliedPluginSnapshot?.snapshotId ?? null, // 应用插件快照ID
          research: meta?.research, // 研究配置
          mediaExecution: mediaExecutionPolicyForProjectMetadata(project.metadata), // 媒体执行策略
          model: choice?.model ?? null, // 选择的模型（如果存在）
          reasoning: choice?.reasoning ?? null, // 推理配置（如果存在）
          titleGeneration: isFirstTurn ? { enabled: true } : undefined, // 首轮对话启用标题生成
          locale, // 语言区域设置
          ...(runAnalyticsHints ? { analyticsHints: runAnalyticsHints } : {}), // 合并运行分析提示（如果存在）
          
          // 运行创建回调：当守护进程创建运行并返回runId时触发
          onRunCreated: (runId) => {
            // 创建固定的助手消息对象，包含runId和排队状态
            const pinnedAssistant = {
              ...latestAssistantMsg, // 复制最新的助手消息
              runId, // 设置运行ID
              runStatus: 'queued' as const, // 标记状态为排队中
            };
            latestAssistantMsg = pinnedAssistant; // 更新最新的助手消息引用
            
            // 视图可能已经切换到不同的项目/会话；
            // 将守护进程运行固定到原始行，以便返回时可以重新连接
            void saveMessage(project.id, runConversationId, pinnedAssistant); // 保存消息到持久化存储
            updateMessageById(assistantId, (prev) => ({ ...prev, runId, runStatus: 'queued' })); // 更新消息状态
          },
          
          // 运行状态变更回调：当守护进程报告运行状态变化时触发
          onRunStatus: (runStatus) => {
            // 如果是终止状态，记录结束时间
            const endedAt = isTerminalRunStatus(runStatus) ? Date.now() : undefined;
            
            // 检查运行是否可以最终化（未被取代的运行才允许）
            const runMayFinalize =
              !supersededRunsRef.current.has(controller);
            
            // 更新助手消息的运行状态和结束时间
            updateMessageById(
              assistantId,
              (prev) => ({
                ...prev,
                runStatus, // 设置运行状态
                endedAt: endedAt === undefined ? prev.endedAt : prev.endedAt ?? endedAt, // 保留已有结束时间，否则使用新的
              }),
              true, // 持久化更新
              runStatus === 'canceled' ? { telemetryFinalized: true } : undefined, // 取消状态时标记遥测已最终化
            );
            
            // 如果运行不可最终化，直接返回
            if (!runMayFinalize) return;
            
            // 更新会话的最新运行状态
            updateConversationLatestRun(runStatus, endedAt);
            
            // 如果是终止运行状态，清除流标记并安排消息刷新
            if (isTerminalRunStatus(runStatus)) {
              clearCurrentRunStreamingMarker(runConversationId, controller, cancelController); // 清除当前运行的流标记
              scheduleConversationMessageRefresh(runConversationId); // 安排会话消息刷新
            }
          },
          
          // 运行事件ID回调：当收到最后一个运行事件ID时触发
          onRunEventId: (lastRunEventId) => {
            updateMessageById(assistantId, (prev) => ({ ...prev, lastRunEventId })); // 更新助手消息的最后事件ID
            persistAssistantSoon(); // 延迟持久化助手消息
          },
        });
        return true; // 返回true表示已处理守护进程模式
      } else {
        // 为非守护进程（BYOK自带密钥）聊天镜像守护进程的聊天路由内存钩子。
        // CLI路径在编写提示之前运行extractFromMessage（因此本轮用户消息中的显式"remember: X"/"我是X"标记
        // 能及时进入本轮系统提示的内存），然后在子进程关闭时排队extractWithLLM
        // （以便小模型从完整的用户+助手交换中提取隐含事实）。
        // BYOK聊天永远不会触发该路径，因此我们在这里针对/api/memory/extract复制两个阶段。
        // 没有这个，即使UI为该模式保存了模型+索引+条目，Memory标签页/模型选择器对BYOK用户也是无效的
        const userText = (userMsg.content ?? '').trim(); // 获取用户消息文本并去除首尾空格
        
        // 快照当前的BYOK聊天配置，以便守护进程可以针对用户正在聊天的相同供应商/密钥/baseUrl/apiVersion
        // 运行"与聊天相同"的内存提取。守护进程本身从不持久化BYOK凭证，因此此每次调用的信号是
        // pickProvider()在没有设置显式内存模型覆盖时避免回退到环境/媒体配置（这对BYOK来说是错误的）的唯一方式。
        // 当聊天配置变化时，选择器会重新同步显式覆盖；此快照覆盖隐式的"与聊天相同"默认值
        const byokChatProvider =
          config.apiProtocol && config.apiKey // 如果配置了API协议和密钥
            ? {
                provider: config.apiProtocol, // API协议类型
                apiKey: config.apiKey, // API密钥
                baseUrl: config.baseUrl, // 基础URL
                apiVersion:
                  config.apiProtocol === 'azure'
                    ? config.apiVersion ?? '' // Azure协议需要API版本
                    : '', // 其他协议不需要
              }
            : undefined; // 未配置则返回undefined
        
        // 如果用户消息不为空，在对话开始前提取记忆
        if (userText.length > 0) {
          try {
            await fetch('/api/memory/extract', {
              method: 'POST', // HTTP POST请求
              headers: { 'Content-Type': 'application/json' }, // 设置内容类型为JSON
              body: JSON.stringify({
                userMessage: userText, // 用户消息文本
                projectId: project.id, // 项目ID
                conversationId: runConversationId, // 会话ID
                chatProvider: byokChatProvider, // 聊天提供者配置
              }),
            });
          } catch {
            // 尽力而为：内存提取绝不应阻塞聊天。
            // 守护进程的SSE总线将在下一个事件时赶上Memory标签页
          }
        }
        
        // 获取组合后的系统提示词
        const systemPrompt = await composedSystemPrompt(runSessionMode);
        
        // 构建包含附件上下文的API历史消息列表
        const apiHistory = await historyWithApiAttachmentContext(
          historyWithCommentAttachmentContext(
            historyWithWorkspaceContext(nextHistory, userMsg.id, runContext), // 添加工区上下文
            userMsg.id, // 用户消息ID
          ),
          userMsg.id, // 用户消息ID
          project.id, // 项目ID
          projectFiles, // 项目文件列表
          { omitNativeImageAttachments: usesAnthropicProxy(config) }, // 使用Anthropic代理时省略原生图片附件
        );
        
        // 推送请求状态事件，显示正在请求的模型
        pushEvent({ kind: 'status', label: 'requesting', detail: config.model });
        // BYOK runs stream client-side and never reach the daemon, so the
        // daemon's authoritative run_created/run_finished are never emitted for
        // them. Emit them here so BYOK runs are counted in the run funnel; the
        // `runtime_type='byok'` rides on these events from the registered
        // super-property. The run id is client-generated (there is no daemon
        // run record). See analytics/byok-run.ts.
        const byokRunId = randomUUID();
        const byokRunBase = {
          projectId: project.id,
          conversationId: runConversationId,
          runId: byokRunId,
          projectKind: null,
          hasAttachment: runAttachments.length > 0,
          userQueryTokens: userText.length > 0 ? Math.ceil(userText.length / 4) : 0,
          model: config.model,
          apiProtocol: config.apiProtocol,
          skillId: project.skillId ?? null,
          sessionMode: (runSessionMode === 'design' ? 'design' : 'ask') as
            | 'design'
            | 'ask',
        };
        trackRunCreated(analytics.track, buildByokRunCreatedProps(byokRunBase));
        const byokRunStartedAt = startedAt;
        let accumulatedAssistantText = '';
        const emitByokRunFinished = (
          result: 'success' | 'failed' | 'cancelled',
          artifactCount: number,
        ): void => {
          trackRunFinished(
            analytics.track,
            buildByokRunFinishedProps({
              ...byokRunBase,
              result,
              artifactCount,
              askedUserQuestion: accumulatedAssistantText.includes('<question-form'),
              totalDurationMs: Math.max(0, Date.now() - byokRunStartedAt),
            }),
          );
        };
        
        // 通过流式传输发送消息
        void streamMessage(config, systemPrompt, apiHistory, controller.signal, {
          // 文本增量回调
          onDelta: (delta) => {
            accumulatedAssistantText += delta; // 累积助手文本
            handlers.onDelta(delta); // 调用通用增量处理器
            handlers.onAgentEvent({ kind: 'text', text: delta }); // 触发文本事件
          },
          
          // 完成回调
          onDone: () => {
            handlers.onDone();
            // Count artifacts produced this turn from the project file diff,
            // mirroring the daemon's run_finished artifact_count. The
            // artifact-count refresh is best-effort: a rejected refetch must
            // NOT swallow run_finished, or a successful BYOK turn leaves the
            // funnel hanging at run_created — the exact gap this path closes.
            void (async () => {
              let artifactCount = 0;
              try {
                const files = await refreshProjectFiles();
                artifactCount = (computeProducedFiles(beforeFileNames, files) ?? []).filter(
                  (f) => Boolean(f.artifactManifest),
                ).length;
              } catch {
                // Refresh failed — still emit run_finished with a 0 count.
              }
              emitByokRunFinished('success', artifactCount);
            })();
            const assistantText = accumulatedAssistantText.trim();
            if (userText.length === 0 || assistantText.length === 0) return;
            void fetch('/api/memory/extract', {
              method: 'POST', // HTTP POST请求
              headers: { 'Content-Type': 'application/json' }, // 设置内容类型为JSON
              body: JSON.stringify({
                userMessage: userText, // 用户消息
                assistantMessage: accumulatedAssistantText, // 助手完整响应
                projectId: project.id, // 项目ID
                conversationId: runConversationId, // 会话ID
                chatProvider: byokChatProvider, // 聊天提供者配置
              }),
            }).catch(() => {
              // 尽力而为：参见上文关于轮次前调用的注释
            });
          },
          onError: (err: Error) => {
            handlers.onError(err);
            emitByokRunFinished(controller.signal.aborted ? 'cancelled' : 'failed', 0);
          },
        }, {
          projectId: project.id, // 项目ID
          // SenseAudio BYOK聊天读取此字段以预填充工具参数的默认模型。
          // 优先使用编辑器中的实时覆盖值；当编辑器下拉菜单选择"使用默认值"时回退到设置中的默认值。
          // 其他协议会忽略未知的请求体字段
          byokImageModel:
            byokImageModelOverride || config.byokImageModel || byokImageModelOptionsPV[0]?.id, // BYOK图像模型选择
          byokVideoModel:
            byokVideoModelOverride || config.byokVideoModel || byokVideoModelOptionsPV[0]?.id, // BYOK视频模型选择
          byokSpeechModel:
            byokSpeechModelOverride || config.byokSpeechModel || byokSpeechModelOptionsPV[0]?.id, // BYOK语音模型选择
          byokSpeechVoice: byokSpeechVoiceOverride || config.byokSpeechVoice, // BYOK语音音色选择
        });
        return true; // 返回true表示已处理API模式
      }
    },
    [
      // 依赖项数组：当这些值变化时，useCallback会重新创建handleSend函数
      attachedComments, // 附加评论列表
      activeConversationId, // 活跃会话ID
      activeSessionMode, // 活跃会话模式
      currentConversationBusy, // 当前会话忙碌状态
      queueChatSendForCurrentConversation, // 排队发送消息的函数
      messages, // 消息列表
      config, // 配置对象（包含模式、API密钥、模型等）
      locale, // 语言区域设置
      agentsById, // 代理ID到代理对象的映射
      // 每个会话的BYOK图像/视频模型覆盖值在此回调内部读取（参见下面的streamMessage上下文）。
      // 如果它们不在依赖项中，下拉菜单会更新其状态和显示，但handleSend保留过时的闭包并发送先前选择的模型
      byokImageModelOverride, // BYOK图像模型覆盖值
      byokVideoModelOverride, // BYOK视频模型覆盖值
      byokSpeechModelOverride, // BYOK语音模型覆盖值
      byokSpeechVoiceOverride, // BYOK语音音色覆盖值
      byokImageModelOptionsPV, // BYOK图像模型选项（持久化值）
      byokVideoModelOptionsPV, // BYOK视频模型选项（持久化值）
      byokSpeechModelOptionsPV, // BYOK语音模型选项（持久化值）
      composedSystemPrompt, // 组合系统提示词的函数
      onTouchProject, // 项目触摸事件处理函数
      project.id, // 项目ID
      project.name, // 项目名称
      projectFiles, // 项目文件列表
      refreshProjectFiles, // 刷新项目文件的函数
      refreshLiveArtifacts, // 刷新实时工件的函数
      readProjectHtml, // 读取项目HTML的函数
      requestOpenFile, // 请求打开文件的函数
      persistMessage, // 持久化消息的函数
      persistMessageById, // 按ID持久化消息的函数
      auditDesignSystemWorkspaceAfterRun, // 运行后审计设计系统工作区的函数
      patchAttachedStatuses, // 更新附件状态的函数
      updateMessageById, // 按ID更新消息的函数
      markStreamingConversation, // 标记流式传输会话的函数
      clearStreamingMarker, // 清除流标记的函数
      clearCurrentRunStreamingMarker, // 清除当前运行流标记的函数
      clearProjectTimeout, // 清除项目超时的函数
      scheduleConversationMessageRefresh, // 安排会话消息刷新的函数
      scheduleProjectTimeout, // 安排项目超时的函数
      onProjectsRefresh, // 项目刷新事件处理函数
      onProjectChange, // 项目变更事件处理函数
    ],
  );

  // Cancel every in-flight run for the current conversation (the user's own
  // streaming turn plus any reattached runs), mark their assistant messages
  // canceled, and drop the streaming state. Defined here — ahead of the
  // queued-send handlers — because "send now" interrupts the active run to
  // make room for the prioritized send.
  const handleStop = useCallback(() => {
    const stoppedAt = Date.now();
    cancelSendTextBuffer(true);
    cancelReattachTextBuffers(true);
    cancelRef.current?.abort();
    cancelRef.current = null;
    for (const controller of reattachCancelControllersRef.current.values()) {
      controller.abort();
    }
    reattachCancelControllersRef.current.clear();
    abortRef.current?.abort();
    abortRef.current = null;
    for (const controller of reattachControllersRef.current.values()) {
      controller.abort();
    }
    reattachControllersRef.current.clear();
    setStreaming(false);
    streamingConversationIdRef.current = null;
    setStreamingConversationId(null);
    setMessages((curr) => {
      const { messages: next, finalized } = finalizeActiveAssistantMessagesOnStop(curr, stoppedAt);
      for (const message of finalized) persistMessage(message, { telemetryFinalized: true });
      return next;
    });
  }, [cancelSendTextBuffer, cancelReattachTextBuffers, persistMessage]);

  // Flip the deck preview to the slide a queued send's marked element lives on
  // the moment that send starts processing. No-op for plain prompts or marks
  // without a slide index; FileWorkspace/FileViewer ignore it unless the named
  // file is the open deck.
  const armSlideNavForQueuedSend = useCallback((item: QueuedChatSend) => {
    const target = queuedSlideNavTarget(item.commentAttachments);
    if (!target) return;
    setSlideNavRequest({ name: target.filePath, slideIndex: target.slideIndex, nonce: Date.now() });
  }, []);

  const sendQueuedChatSendNow = useCallback((id: string) => {
    const item = queuedChatSendsRef.current.find((candidate) => candidate.id === id);
    if (!item) return;
    if (currentConversationBusy) {
      // "Send now" while the agent is still working: the user has explicitly
      // chosen this turn over the in-flight one, so interrupt the running run
      // and move this item to the front. Stopping flips the conversation out
      // of its busy state, and the auto-start effect below then flushes the
      // now-first queued send — reusing the same path as a natural completion,
      // so runs never overlap.
      //
      // Record the runs we're superseding BEFORE handleStop() clears the active
      // refs. The daemon still delivers a late terminal callback for the
      // canceled run; tagging its controller here lets those callbacks be
      // recognized as stale and skip every current-run side effect, even if the
      // replacement send hasn't attached yet.
      if (abortRef.current) supersededRunsRef.current.add(abortRef.current);
      for (const controller of reattachControllersRef.current.values()) {
        supersededRunsRef.current.add(controller);
      }
      // The interrupted turn moved its preview-comment attachments to
      // 'applying' when it started; since we now suppress its terminal
      // callbacks, reset them to 'open' so they don't stay stuck mid-apply.
      // Reset ONLY the in-flight run's comments: queued sends (including the
      // one being prioritized) also hold their attachments in 'applying', and
      // those must stay reserved — the replacement run re-applies them. The
      // in-flight run's comments are exactly the 'applying' ones not owned by
      // any queued send.
      const queuedCommentIds = new Set(
        queuedChatSendsRef.current.flatMap((send) =>
          send.commentAttachments.map((attachment) => attachment.id),
        ),
      );
      const stuckApplying = previewCommentsRef.current.filter(
        (comment) => comment.status === 'applying' && !queuedCommentIds.has(comment.id),
      );
      if (stuckApplying.length > 0) {
        const resetIds = new Set(stuckApplying.map((comment) => comment.id));
        setPreviewComments((current) =>
          current.map((comment) =>
            resetIds.has(comment.id) ? { ...comment, status: 'open' } : comment,
          ),
        );
        void Promise.all(
          stuckApplying.map((comment) =>
            patchPreviewCommentStatus(project.id, comment.conversationId, comment.id, 'open'),
          ),
        ).catch(() => {});
      }
      prioritizeQueuedChatSend(id);
      handleStop();
      return;
    }
    void (async () => {
      armSlideNavForQueuedSend(item);
      const started = await handleSend(
        item.prompt,
        item.attachments,
        item.commentAttachments,
        item.meta,
      );
      if (started) removeQueuedChatSend(id);
    })();
  }, [armSlideNavForQueuedSend, currentConversationBusy, handleSend, handleStop, prioritizeQueuedChatSend, project.id, removeQueuedChatSend]);

  useEffect(() => {
    if (currentConversationBusy) {
      startingQueuedChatSendIdRef.current = null;
      return;
    }
    if (startingQueuedChatSendIdRef.current) return;
    if (!activeConversationId) return;
    if (messagesConversationIdRef.current !== activeConversationId) return;
    const next = queuedChatSendsRef.current.find(
      (item) => item.conversationId === activeConversationId,
    );
    if (!next) return;
    startingQueuedChatSendIdRef.current = next.id;
    armSlideNavForQueuedSend(next);
    void (async () => {
      const started = await handleSend(
        next.prompt,
        next.attachments,
        next.commentAttachments,
        next.meta,
      );
      if (!started) {
        if (startingQueuedChatSendIdRef.current === next.id) {
          startingQueuedChatSendIdRef.current = null;
        }
        return;
      }
      removeQueuedChatSend(next.id);
      scheduleProjectTimeout(() => {
        if (startingQueuedChatSendIdRef.current !== next.id) return;
        startingQueuedChatSendIdRef.current = null;
        setQueuedAutoStartTick((tick) => tick + 1);
      }, 0);
    })();
  }, [
    activeConversationId,
    armSlideNavForQueuedSend,
    currentConversationBusy,
    queuedAutoStartTick,
    queuedChatSends,
    handleSend,
    removeQueuedChatSend,
    scheduleProjectTimeout,
  ]);

  const handleRetry = useCallback(
    (assistantMessage: ChatMessage) => {
      if (currentConversationActionDisabled) return;
      void handleSend('', [], [], { retryOfAssistantId: assistantMessage.id });
    },
    [currentConversationActionDisabled, handleSend],
  );

  // "Continue" on a resumable failed run: send a fresh turn in the same
  // conversation. For a session-resuming runtime (Claude) the daemon persisted
  // the failed run's CLI session, so this turn resumes it (`--resume`) and the
  // agent continues from its committed work instead of restarting. Mirrors the
  // "Continue remaining tasks" affordance; unlike Retry it does not replay the
  // prior turn from scratch. Tagged `entryFrom: 'resume_continue'` so
  // run_created / run_finished can quantify how often resume fires and whether
  // it recovers (the whole point is to show the mechanism lowers failure rate).
  const handleResumeRun = useCallback(
    (_assistantMessage: ChatMessage) => {
      if (currentConversationActionDisabled) return;
      void handleSend(RESUME_CONTINUE_PROMPT, [], [], { entryFrom: 'resume_continue' });
    },
    [currentConversationActionDisabled, handleSend],
  );

  // "Switch to AMR & retry" from the failed-run card: switch the run to AMR,
  // open Settings on the AMR controls so the user can sign in / authorize /
  // top up, and arm an auto-retry that fires once AMR is selected AND signed
  // in (see the effect below).
  const [pendingAmrRetry, setPendingAmrRetry] = useState<ChatMessage | null>(null);
  const handleSwitchToAmrAndRetry = useCallback(
    (failedAssistant: ChatMessage) => {
      if (currentConversationActionDisabled) return;
      onModeChange('daemon');
      onAgentChange('amr');
      onOpenAmrSettings?.();
      setPendingAmrRetry(failedAssistant);
    },
    [currentConversationActionDisabled, onModeChange, onAgentChange, onOpenAmrSettings],
  );
  // PR #3157: Antigravity's `agy -p` cannot complete OAuth on its own,
  // so the auth banner offers a one-click "Sign in via terminal"
  // button that POSTs to the daemon. The daemon opens a system
  // Terminal running `agy` (osascript / x-terminal-emulator /
  // `cmd /c start`); the user finishes Google sign-in there and then
  // clicks Retry to redo the chat run. We don't auto-retry because
  // the OAuth completion happens externally with no reliable signal
  // back to the chat — the secondary Retry button on the same banner
  // covers the manual case.
  const handleLaunchAntigravityOauth = useCallback(async () => {
    try {
      const { launchAntigravityOauth } = await import('../providers/daemon');
      const result = await launchAntigravityOauth();
      if (!result.ok) {
        // Surface the daemon-side reason so the user knows whether
        // the spawn failed because of missing osascript / unsupported
        // platform / etc. instead of silently swallowing it.
        console.warn('[antigravity] oauth-launch failed:', result.error);
      }
    } catch (err) {
      console.warn('[antigravity] oauth-launch threw:', err);
    }
  }, []);
  // Poll the AMR login status while a retry is armed, rather than only reacting
  // to the AmrLoginPill's status event — the user may close Settings (which
  // unmounts the pill and stops its polling) before finishing sign-in in the
  // browser. Polling here keeps working regardless of the pill's lifecycle.
  // Fires once AMR is the selected agent AND the account is signed in.
  useEffect(() => {
    if (!pendingAmrRetry) return;
    let cancelled = false;
    const tryRetry = async () => {
      if (cancelled) return;
      if (!(config.mode === 'daemon' && config.agentId === 'amr')) return;
      const status = await fetchVelaLoginStatus().catch(() => null);
      if (cancelled || status?.loggedIn !== true) return;
      setPendingAmrRetry(null);
      handleRetry(pendingAmrRetry);
    };
    void tryRetry();
    const interval = setInterval(() => void tryRetry(), 2000);
    // Give up after a few minutes so we never poll forever.
    const stop = setTimeout(() => {
      if (!cancelled) setPendingAmrRetry(null);
    }, 5 * 60 * 1000);
    return () => {
      cancelled = true;
      clearInterval(interval);
      clearTimeout(stop);
    };
  }, [pendingAmrRetry, config.mode, config.agentId, handleRetry]);

  useEffect(() => {
    if (!autoAuditRepairSeed) return;
    if (!activeConversationId) return;
    if (!messagesInitialized) return;
    if (currentConversationBusy) return;
    const repairText = autoAuditRepairSeed.value.trim();
    setAutoAuditRepairSeed(null);
    if (!repairText) return;
    void handleSend(repairText, [], []);
  }, [
    activeConversationId,
    autoAuditRepairSeed,
    currentConversationBusy,
    handleSend,
    messagesInitialized,
  ]);

  const handleSendBoardCommentAttachments = useCallback(
    async (commentAttachments: ChatCommentAttachment[], images: File[] = []) => {
      if (currentConversationQueueDisabled) return false;
      if (commentAttachments.length === 0 && images.length === 0) return false;
      setWorkspaceFocused(false);
      setCommentInspectorActive(false);
      // Upload any attached images once, then queue. Each comment becomes its
      // own task (so multiple notes => multiple queued tasks); the images ride
      // along the first task rather than being duplicated across every note.
      let uploaded: ChatAttachment[] = [];
      if (images.length > 0) {
        const result = await uploadProjectFiles(project.id, images);
        uploaded = result.uploaded;
      }
      if (commentAttachments.length === 0) {
        if (uploaded.length > 0) await handleSend('', uploaded, [], { queueOnly: true, entryFrom: 'comment' });
        return true;
      }
      for (let i = 0; i < commentAttachments.length; i++) {
        const commentAttachment = commentAttachments[i]!;
        const savedImages = chatAttachmentsFromPreviewCommentImages(commentAttachment.imageAttachments);
        const prompt = commentTaskQuery(commentAttachment);
        // Comment/board pin → run: tag entry_from='comment' so the dashboard
        // separates annotation-driven runs from plain composer sends.
        await handleSend(
          prompt,
          mergeChatAttachments(i === 0 ? uploaded : [], savedImages),
          [commentTaskContextAttachment(commentAttachment)],
          { queueOnly: true, entryFrom: 'comment' },
        );
      }
      return true;
    },
    [handleSend, project.id, currentConversationQueueDisabled],
  );
  const commentQueueOnSend = currentConversationBusy && !currentConversationQueueDisabled;

  const handleContinueRemainingTasks = useCallback(
    (_assistantMessage: ChatMessage, todos: TodoItem[]) => {
      if (currentConversationActionDisabled || todos.length === 0) return;
      const remainingList = todos
        .map((todo, i) => {
          const label =
            todo.status === 'in_progress' && todo.activeForm ? todo.activeForm : todo.content;
          return `${i + 1}. [${todo.status}] ${label}`;
        })
        .join('\n');
      const prompt =
        'Continue the remaining unfinished tasks from the previous run. ' +
        'Do not redo completed work. Focus only on these unfinished todos:\n\n' +
        `${remainingList}\n\n` +
        'Before making changes, inspect the current project files as needed. ' +
        'Update TodoWrite as you complete each remaining task.';
      void handleSend(prompt, [], []);
    },
    [currentConversationActionDisabled, handleSend],
  );

  const selectedPluginActionAgent =
    config.mode === 'daemon' && config.agentId
      ? agentsById.get(config.agentId)
      : null;
  const selectedPluginActionChoice =
    config.mode === 'daemon' && config.agentId
      ? config.agentModels?.[config.agentId]
      : undefined;
  const effectiveSelectedPluginActionChoice = effectiveAgentModelChoice(
    selectedPluginActionAgent,
    selectedPluginActionChoice,
  );
  const pluginWorkflowAgentName =
    config.mode === 'daemon'
      ? agentModelDisplayName(
          config.agentId,
          selectedPluginActionAgent?.name,
          effectiveSelectedPluginActionChoice?.model,
        )
      : apiProtocolModelLabel(config.apiProtocol, config.model);

  const handlePluginFolderAgentAction = useCallback(
    async (relativePath: string, action: PluginFolderAgentAction) => {
      if (currentConversationActionDisabled || !activeConversationId) return;
      setHiddenAssistantPluginActionPaths((prev) => new Set(prev).add(relativePath));
      if (action === 'install') {
        setActivePluginActionPaths((prev) => new Set(prev).add(relativePath));
        let outcome;
        try {
          outcome = await installGeneratedPluginFolder(project.id, relativePath);
        } finally {
          setActivePluginActionPaths((prev) => {
            const next = new Set(prev);
            next.delete(relativePath);
            return next;
          });
          setHiddenAssistantPluginActionPaths((prev) => {
            const next = new Set(prev);
            next.delete(relativePath);
            return next;
          });
        }
        if (!outcome.ok) throw new Error(outcome.message);
        return { message: outcome.message };
      }
      const conversationId = activeConversationId;
      const shareAction = action === 'publish' ? 'publish-github' : 'contribute-open-design';
      setActivePluginActionPaths((prev) => new Set(prev).add(relativePath));
      let taskStart;
      try {
        taskStart = await startGeneratedPluginShareTask(project.id, relativePath, shareAction);
      } catch (error) {
        setActivePluginActionPaths((prev) => {
          const next = new Set(prev);
          next.delete(relativePath);
          return next;
        });
        setHiddenAssistantPluginActionPaths((prev) => {
          const next = new Set(prev);
          next.delete(relativePath);
          return next;
        });
        throw error;
      }
      const startedAt = taskStart.startedAt;
      const messageId = randomUUID();
      const updateConversationLatestRun = (
        status: NonNullable<ChatMessage['runStatus']>,
        endedAt?: number,
      ) => {
        setConversations((curr) =>
          curr.map((conversation) =>
            conversation.id === conversationId
              ? {
                  ...conversation,
                  updatedAt: endedAt ?? startedAt,
                  latestRun: {
                    status,
                    startedAt,
                    ...(endedAt === undefined
                      ? {}
                      : {
                          endedAt,
                          durationMs: Math.max(0, endedAt - startedAt),
                        }),
                  },
                }
              : conversation,
          ),
        );
      };
      const progressMessage: ChatMessage = {
        id: messageId,
        role: 'assistant',
        content: pluginWorkflowStartContent(action, relativePath),
        agentName: pluginWorkflowAgentName,
        events: pluginWorkflowPlannedEvents(action, relativePath),
        createdAt: startedAt,
        startedAt,
        runStatus: 'running',
      };
      setForceStreamingPluginMessageIds((prev) => new Set(prev).add(messageId));
      appendConversationMessage(conversationId, progressMessage, undefined, false);
      updateConversationLatestRun('running');
      void (async () => {
        let since = 0;
        let liveEvents = [...pluginWorkflowPlannedEvents(action, relativePath)];
        let liveContent = pluginWorkflowStartContent(action, relativePath);
        while (true) {
          const snapshot = await waitGeneratedPluginShareTask(taskStart.taskId, since, 25_000);
          since = snapshot.nextSince;
          if (snapshot.progress.length > 0) {
            const newTextEvents = snapshot.progress
              .map((line) => line.trim())
              .filter(Boolean)
              .map((line) => ({ kind: 'text' as const, text: `${line}\n` }));
            liveEvents = [
              ...liveEvents.filter((event, index) => !(index === liveEvents.length - 1 && event.kind === 'status' && event.label === 'working')),
              ...newTextEvents,
              { kind: 'status', label: 'working', detail: pluginWorkflowTitle(action) },
            ];
            liveContent = `${liveContent}\n\n${snapshot.progress.map((line) => line.trim()).filter(Boolean).join('\n')}`.trim();
            replaceConversationMessage(
              conversationId,
              {
                ...progressMessage,
                content: liveContent,
                events: liveEvents,
                runStatus: 'running',
              },
              undefined,
              false,
            );
          }
          if (snapshot.status === 'running' || snapshot.status === 'queued') continue;
          const endedAt = snapshot.endedAt ?? Date.now();
          setActivePluginActionPaths((prev) => {
            const next = new Set(prev);
            next.delete(relativePath);
            return next;
          });
          setHiddenAssistantPluginActionPaths((prev) => {
            const next = new Set(prev);
            next.delete(relativePath);
            return next;
          });
          if (snapshot.status === 'done' && snapshot.result) {
            setForceStreamingPluginMessageIds((prev) => {
              const next = new Set(prev);
              next.delete(messageId);
              return next;
            });
            replaceConversationMessage(
              conversationId,
              {
                ...progressMessage,
                content: pluginWorkflowSuccessContent(
                  action,
                  relativePath,
                  snapshot.result.message,
                  snapshot.result.url,
                  snapshot.result.log,
                ),
                events: pluginWorkflowResultEvents(
                  action,
                  relativePath,
                  snapshot.result.message,
                  snapshot.result.url,
                  snapshot.result.log,
                  true,
                  liveEvents,
                ),
                endedAt,
                runStatus: 'succeeded',
              },
              { telemetryFinalized: true },
            );
            updateConversationLatestRun('succeeded', endedAt);
            return;
          }
          const errorMessage = snapshot.error?.message || `${pluginWorkflowTitle(action)} failed.`;
          setForceStreamingPluginMessageIds((prev) => {
            const next = new Set(prev);
            next.delete(messageId);
            return next;
          });
          replaceConversationMessage(
            conversationId,
            {
              ...progressMessage,
              content: pluginWorkflowFailureContent(
                action,
                relativePath,
                errorMessage,
                snapshot.error?.log,
              ),
              events: pluginWorkflowResultEvents(
                action,
                relativePath,
                errorMessage,
                undefined,
                snapshot.error?.log,
                false,
                liveEvents,
              ),
              endedAt,
              runStatus: 'failed',
            },
            { telemetryFinalized: true },
          );
          updateConversationLatestRun('failed', endedAt);
          return;
        }
      })().catch((err) => {
        const endedAt = Date.now();
        setForceStreamingPluginMessageIds((prev) => {
          const next = new Set(prev);
          next.delete(messageId);
          return next;
        });
        setActivePluginActionPaths((prev) => {
          const next = new Set(prev);
          next.delete(relativePath);
          return next;
        });
        setHiddenAssistantPluginActionPaths((prev) => {
          const next = new Set(prev);
          next.delete(relativePath);
          return next;
        });
        replaceConversationMessage(
          conversationId,
          {
            ...progressMessage,
            content: pluginWorkflowFailureContent(
              action,
              relativePath,
              err instanceof Error ? err.message : String(err),
            ),
            events: pluginWorkflowResultEvents(
              action,
              relativePath,
              err instanceof Error ? err.message : String(err),
              undefined,
              [],
              false,
            ),
            endedAt,
            runStatus: 'failed',
          },
          { telemetryFinalized: true },
        );
        updateConversationLatestRun('failed', endedAt);
      });
      return;
    },
    [
      activeConversationId,
      appendConversationMessage,
      currentConversationActionDisabled,
      pluginWorkflowAgentName,
      project.id,
      replaceConversationMessage,
    ],
  );

  // "Share to Open Design" — kicks off the bundled `od-share-to-community`
  // scenario in the active conversation. We just inject the trigger prompt
  // through the standard chat-send path; the agent then loads SKILL.md and
  // drives the rest. Keep this preparing state alive for the resulting chat
  // run so the action reads as async packaging instead of instant sharing.
  const [shareToOpenDesignBusyMessageId, setShareToOpenDesignBusyMessageId] = useState<string | null>(null);
  const shareToOpenDesignBusyMessageIdRef = useRef<string | null>(null);
  useEffect(() => {
    if (!shareToOpenDesignBusyMessageIdRef.current || currentConversationBusy) return;
    shareToOpenDesignBusyMessageIdRef.current = null;
    setShareToOpenDesignBusyMessageId(null);
  }, [currentConversationBusy]);
  const handleShareToOpenDesign = useCallback((assistantMessageId: string) => {
    if (currentConversationActionDisabled || shareToOpenDesignBusyMessageIdRef.current) return;
    shareToOpenDesignBusyMessageIdRef.current = assistantMessageId;
    setShareToOpenDesignBusyMessageId(assistantMessageId);
    void Promise.resolve(handleSend(SHARE_TO_COMMUNITY_PROMPT, [], []))
      .then((started) => {
        if (started) return;
        shareToOpenDesignBusyMessageIdRef.current = null;
        setShareToOpenDesignBusyMessageId(null);
      })
      .catch(() => {
        shareToOpenDesignBusyMessageIdRef.current = null;
        setShareToOpenDesignBusyMessageId(null);
      });
  }, [currentConversationActionDisabled, handleSend]);

  const sentDesignSystemReviewTaskKeysRef = useRef<Set<string>>(new Set());
  const persistDesignSystemReviewEntry = useCallback((
    sectionTitle: string,
    entry: DesignSystemReviewEntry,
  ) => {
    const baseMetadata: ProjectMetadata = {
      kind: project.metadata?.kind ?? 'other',
      ...project.metadata,
    };
    const metadata: ProjectMetadata = {
      ...baseMetadata,
      designSystemReview: {
        ...(baseMetadata.designSystemReview ?? {}),
        [sectionTitle]: entry,
      },
    };
    onProjectChange({ ...project, metadata });
    void patchProject(project.id, { metadata });
  }, [onProjectChange, project]);

  const sendDesignSystemFeedback = useCallback((
    sectionTitle: string,
    feedback: string,
    sectionFiles: string[],
  ): DesignSystemReviewAgentTask | void => {
    const cleanFeedback = feedback.trim();
    if (!cleanFeedback) return;
    const prompt = designSystemNeedsWorkPrompt(sectionTitle, cleanFeedback, sectionFiles);
    const queuedAt = new Date().toISOString();
    if (!activeConversationId || !messagesInitialized || currentConversationActionDisabled) {
      return {
        status: 'queued',
        prompt,
        queuedAt,
      };
    }
    const task: DesignSystemReviewAgentTask = {
      status: 'sent',
      prompt,
      queuedAt,
      sentAt: queuedAt,
    };
    sentDesignSystemReviewTaskKeysRef.current.add(`${sectionTitle}:${queuedAt}`);
    void handleSend(prompt, designSystemFeedbackAttachments(projectFiles, sectionFiles), []);
    return task;
  }, [
    activeConversationId,
    currentConversationActionDisabled,
    handleSend,
    messagesInitialized,
    projectFiles,
  ]);
  const persistDesignSystemReviewDecision = useCallback((
    sectionTitle: string,
    decision: DesignSystemReviewEntry['decision'],
    details?: DesignSystemReviewDetails,
  ) => {
    const entry: DesignSystemReviewEntry = {
      decision,
      updatedAt: new Date().toISOString(),
    };
    if (details?.feedback) entry.feedback = details.feedback;
    if (details?.files) entry.files = details.files;
    if (details?.agentTask) entry.agentTask = details.agentTask;
    persistDesignSystemReviewEntry(sectionTitle, entry);
  }, [persistDesignSystemReviewEntry]);
  useEffect(() => {
    if (!activeConversationId || !messagesInitialized || currentConversationActionDisabled) return;
    const queued = Object.entries(project.metadata?.designSystemReview ?? {}).find(
      ([, entry]) =>
        entry.decision === 'needs-work'
        && Boolean(entry.feedback?.trim())
        && entry.agentTask?.status === 'queued',
    );
    if (!queued) return;
    const [sectionTitle, entry] = queued;
    const task = entry.agentTask;
    if (!task) return;
    const taskKey = `${sectionTitle}:${task.queuedAt}`;
    if (sentDesignSystemReviewTaskKeysRef.current.has(taskKey)) return;
    sentDesignSystemReviewTaskKeysRef.current.add(taskKey);
    const sectionFiles = entry.files ?? [];
    const prompt = task.prompt || designSystemNeedsWorkPrompt(
      sectionTitle,
      entry.feedback ?? '',
      sectionFiles,
    );
    const sentAt = new Date().toISOString();
    persistDesignSystemReviewEntry(sectionTitle, {
      ...entry,
      agentTask: {
        ...task,
        status: 'sent',
        prompt,
        sentAt,
      },
    });
    void handleSend(prompt, designSystemFeedbackAttachments(projectFiles, sectionFiles), []);
  }, [
    activeConversationId,
    currentConversationActionDisabled,
    handleSend,
    messagesInitialized,
    persistDesignSystemReviewEntry,
    project.metadata?.designSystemReview,
    projectFiles,
  ]);

  const handleExportAsPptx = useCallback(
    (fileName: string) => {
      if (currentConversationActionDisabled) return;
      const prompt = buildPptxExportPrompt(fileName);
      const attachment: ChatAttachment = {
        path: fileName,
        name: fileName,
        kind: 'file',
      };
      void handleSend(prompt, [attachment], []);
    },
    [currentConversationActionDisabled, handleSend],
  );

  const handleNewConversation = useCallback(async () => {
    if (creatingConversationRef.current) return;
    // Only block if we're sure the current conversation is empty:
    // messages must be loaded AND match the active conversation.
    if (
      messagesConversationIdRef.current === activeConversationId &&
      messages.length === 0
    ) {
      return;
    }
    creatingConversationRef.current = true;
    setCreatingConversation(true);
    setConversationLoadError(null);
    try {
      const fresh = await createConversation(project.id);
      if (!fresh) throw new Error('Could not create a conversation for this project.');
      // Eagerly clear messages and update ref so rapid clicks don't create
      // duplicate empty conversations before the effect resolves.
      setMessages([]);
      setStreaming(false);
      streamingConversationIdRef.current = null;
      setStreamingConversationId(null);
      setMessagesConversationId(null);
      messagesConversationIdRef.current = fresh.id;
      setConversations((curr) => [fresh, ...curr]);
      setActiveConversationId(fresh.id);
      // Push the new conversation id into the URL synchronously so the
      // route-sync effect sees a matching `routeConversationId` before
      // it can revert `activeConversationId`. Without this, the route-sync
      // effect can fight the conversation switch, preventing users from
      // switching back to older conversations after creating a new one.
      navigate(
        {
          kind: 'project',
          projectId: project.id,
          conversationId: fresh.id,
          fileName: openTabsState.active ?? null,
        },
        { replace: true },
      );
      setError(null);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Could not create a conversation for this project.';
      setConversationLoadError(message);
      setError(message);
    } finally {
      creatingConversationRef.current = false;
      setCreatingConversation(false);
    }
  }, [project.id, activeConversationId, messages.length, navigate, openTabsState.active]);

  const handleSelectConversation = useCallback((id: string) => {
    if (id === activeConversationId && failedMessagesConversationId !== id) return;
    setMessages([]);
    setPreviewComments([]);
    setAttachedComments([]);
    setArtifact(null);
    setStreaming(false);
    streamingConversationIdRef.current = null;
    setStreamingConversationId(null);
    setMessagesConversationId(null);
    setFailedMessagesConversationId(null);
    setConversationLoadError(null);
    messagesConversationIdRef.current = null;
    setActiveConversationId(id);
    // Push the new conversation id into the URL synchronously so the
    // route-sync effect at L512 sees a matching `routeConversationId`
    // before it can find the previous conversation in the list and
    // revert `activeConversationId` to it. Without this, the same
    // effect that fights handleNewConversation also fights chat
    // switching, ping-ponging until React's nested-update guard fires.
    navigate(
      {
        kind: 'project',
        projectId: project.id,
        conversationId: id,
        fileName: openTabsState.active ?? null,
      },
      { replace: true },
    );
    setMessageLoadRetryNonce((nonce) => nonce + 1);
  }, [activeConversationId, failedMessagesConversationId, project.id, openTabsState.active]);

  const handleDeleteConversation = useCallback(
    async (id: string) => {
      const ok = await deleteConversationApi(project.id, id);
      if (!ok) return;
      // The deleted conversation may have owned an unanswered
      // `<question-form>`, which the daemon counts toward the project's
      // `needsInput` flag in `/api/projects`. Home cards render that
      // flag from the cached projects payload, so without refreshing
      // it here the `Needs input` badge survives the deletion until
      // the next manual reload.
      onProjectsRefresh();
      setConversations((curr) => {
        const next = curr.filter((c) => c.id !== id);
        if (next.length === 0) {
          // Re-seed so the project always has at least one conversation
          // to write into.
          void createConversation(project.id).then((fresh) => {
            if (fresh) {
              setConversations([fresh]);
              setActiveConversationId(fresh.id);
            }
          });
        } else if (id === activeConversationId) {
          setActiveConversationId(next[0]!.id);
        }
        return next;
      });
    },
    [project.id, activeConversationId, onProjectsRefresh],
  );

  const handleRenameConversation = useCallback(
    async (id: string, title: string) => {
      const trimmed = title.trim() || null;
      setConversations((curr) =>
        curr.map((c) => (c.id === id ? { ...c, title: trimmed } : c)),
      );
      await patchConversation(project.id, id, { title: trimmed });
    },
    [project.id],
  );

  const handleConversationSessionModeChange = useCallback(
    async (id: string, sessionMode: ChatSessionMode) => {
      setConversations((curr) =>
        curr.map((conversation) =>
          conversation.id === id ? { ...conversation, sessionMode } : conversation,
        ),
      );
      const updated = await patchConversation(project.id, id, { sessionMode });
      if (updated) {
        setConversations((curr) =>
          curr.map((conversation) =>
            conversation.id === id ? { ...conversation, ...updated } : conversation,
          ),
        );
      }
    },
    [project.id],
  );

  const handleActiveConversationSessionModeChange = useCallback(
    (sessionMode: ChatSessionMode) => {
      if (!activeConversationId) return;
      void handleConversationSessionModeChange(activeConversationId, sessionMode);
    },
    [activeConversationId, handleConversationSessionModeChange],
  );

  const handleForkFromMessage = useCallback(
    async (assistantMessage: ChatMessage) => {
      if (!activeConversationId || forkingMessageId) return;
      setForkingMessageId(assistantMessage.id);
      setConversationLoadError(null);
      try {
        const sourceTitle = activeConversation?.title?.trim();
        const forkTitle = sourceTitle
          ? t('chat.forkedConversationTitle', { title: sourceTitle })
          : undefined;
        // Seed the fork from the messages the user is actually looking at,
        // up to and including the fork point. A run that errored or had its
        // connection reset before its assistant message was persisted leaves
        // that message in memory only; copying from the database by id would
        // 404 and silently drop the fork. Sending the in-memory snapshot makes
        // the fork resilient to that gap.
        const forkIndex = messages.findIndex((m) => m.id === assistantMessage.id);
        const seedMessages =
          forkIndex >= 0 ? messages.slice(0, forkIndex + 1) : [...messages, assistantMessage];
        const fresh = await createConversation(project.id, forkTitle, {
          seedFromConversationId: activeConversationId,
          forkAfterMessageId: assistantMessage.id,
          sessionMode: activeSessionMode,
          seedMessages,
        });
        if (!fresh) throw new Error(t('chat.forkConversationFailed'));
        setMessages([]);
        setPreviewComments([]);
        setAttachedComments([]);
        setArtifact(null);
        setStreaming(false);
        streamingConversationIdRef.current = null;
        setStreamingConversationId(null);
        setMessagesConversationId(null);
        messagesConversationIdRef.current = null;
        setFailedMessagesConversationId(null);
        setConversations((curr) => [fresh, ...curr.filter((c) => c.id !== fresh.id)]);
        setActiveConversationId(fresh.id);
        navigate(
          {
            kind: 'project',
            projectId: project.id,
            conversationId: fresh.id,
            fileName: openTabsState.active ?? null,
          },
          { replace: true },
        );
        onProjectsRefresh();
        setError(null);
      } catch (err) {
        const message = err instanceof Error ? err.message : t('chat.forkConversationFailed');
        setConversationLoadError(message);
        setError(message);
      } finally {
        setForkingMessageId(null);
      }
    },
    [
      activeConversationId,
      activeConversation?.title,
      activeSessionMode,
      forkingMessageId,
      messages,
      navigate,
      onProjectsRefresh,
      openTabsState.active,
      project.id,
      t,
    ],
  );

  const handleProjectRename = useCallback(
    (newName: string) => {
      const trimmed = newName.trim();
      if (!trimmed || trimmed === project.name) return;
      const metadata = project.metadata
        ? { ...project.metadata, nameSource: 'user' as const }
        : undefined;
      const updated: Project = {
        ...project,
        name: trimmed,
        ...(metadata ? { metadata } : {}),
        updatedAt: Date.now(),
      };
      onProjectChange(updated);
      void patchProject(project.id, {
        name: trimmed,
        ...(metadata ? { metadata } : {}),
      });
    },
    [project, onProjectChange],
  );

  const activeConversationChatState = useMemo(
    () =>
      activeConversationId
        ? {
	            conversationId: activeConversationId,
	            messages,
	            streaming: currentConversationStreaming,
	            loading: currentConversationLoading,
	            sendDisabled: currentConversationSendDisabled,
            queuedItems: currentConversationQueuedItems,
            error: conversationLoadError ?? error ?? audioVoiceOptionsError,
            onSend: handleSend,
            onRetry: handleRetry,
            onStop: handleStop,
            onRemoveQueuedSend: removeQueuedChatSend,
            onUpdateQueuedSend: updateQueuedChatSend,
            onReorderQueuedSends: reorderCurrentConversationQueuedChatSends,
            onSendQueuedNow: sendQueuedChatSendNow,
            onAssistantFeedback: handleAssistantFeedback,
          }
        : undefined,
    [
      activeConversationId,
      audioVoiceOptionsError,
      conversationLoadError,
      currentConversationActionDisabled,
	      currentConversationQueuedItems,
	      currentConversationSendDisabled,
	      currentConversationLoading,
	      currentConversationStreaming,
      error,
      handleAssistantFeedback,
      handleRetry,
      handleSend,
      handleStop,
      messages,
      removeQueuedChatSend,
      reorderCurrentConversationQueuedChatSends,
      sendQueuedChatSendNow,
      updateQueuedChatSend,
    ],
  );

  const handleChangeDesignSystemId = useCallback(
    (nextId: string | null) => {
      if ((project.designSystemId ?? null) === nextId) return;
      // `design_system_apply_result` studio variant. The existing
      // NewProjectPanel picker fires the same event under
      // `page_name=home`; this in-project header picker fires under
      // `page_name=studio` so the funnel sees applies from both
      // surfaces. `target_project_kind` derives from
      // `project.metadata.kind`.
      const target =
        (projectKindToTracking(project.metadata?.kind ?? null, project.metadata?.videoModel) ?? 'unknown') as TrackingDesignSystemApplyTargetKind;
      const picked = nextId
        ? designSystems.find((d) => d.id === nextId)
        : null;
      const origin: TrackingDesignSystemOrigin | undefined = picked
        ? picked.source === 'user'
          ? 'manual_create'
          : picked.source === 'built-in'
            ? 'official_preset'
            : picked.source === 'installed'
              ? 'template'
              : 'unknown'
        : undefined;
      const status: TrackingDesignSystemStatusValue | undefined = picked
        ? picked.status === 'draft' || picked.status === 'published'
          ? picked.status
          : 'unknown'
        : undefined;
      if (nextId === null) {
        trackDesignSystemApplyResult(analytics.track, {
          page_name: 'studio',
          area: 'design_system_picker',
          action: 'clear_selection',
          result: 'success',
          target_project_kind: target,
          design_system_applied: false,
          design_system_selection_mode: 'none',
          is_default: false,
          is_auto_selected: false,
          available_design_system_count: designSystems.length,
          duration_ms: 0,
        });
      } else {
        trackDesignSystemApplyResult(analytics.track, {
          page_name: 'studio',
          area: 'design_system_picker',
          action: 'select_design_system',
          result: 'success',
          target_project_kind: target,
          design_system_id: nextId,
          design_system_source: origin,
          design_system_status: status,
          design_system_applied: true,
          design_system_selection_mode: 'manual',
          is_default: false,
          is_auto_selected: false,
          available_design_system_count: designSystems.length,
          duration_ms: 0,
        });
      }
      const updated: Project = {
        ...project,
        designSystemId: nextId,
        updatedAt: Date.now(),
      };
      onProjectChange(updated);
      void patchProject(project.id, { designSystemId: nextId });
    },
    [project, onProjectChange, designSystems, analytics.track],
  );

  const projectMeta = useMemo(() => {
    // Design system is rendered by the adjacent picker chip — keep the
    // bare meta string focused on skill / mode so the two surfaces
    // don't show the same label twice.
    const summary =
      skills.find((s) => s.id === project.skillId) ??
      designTemplates.find((s) => s.id === project.skillId);
    const skill = summary?.name;
    return skill ?? t('project.metaFreeform');
  }, [skills, designTemplates, project.skillId, t]);

  const activeDesignSystemSummary = useMemo(() => {
    if (!project.designSystemId) return null;
    return designSystems.find((d) => d.id === project.designSystemId) ?? null;
  }, [designSystems, project.designSystemId]);

  const designSystemProject = useMemo(() => {
    if (project.metadata?.importedFrom !== 'design-system') return null;
    if (!project.designSystemId) return null;
    return designSystems.find((d) => d.id === project.designSystemId) ?? null;
  }, [designSystems, project.designSystemId, project.metadata?.importedFrom]);
  const designSystemActivityEvents = useMemo(
    () => designSystemProject ? latestDesignSystemActivityEvents(messages) : [],
    [designSystemProject, messages],
  );
  const connectRepoNeeded = useMemo(
    () => designSystemNeedsRepoConnect(designSystemProject, projectFiles.map((file) => file.name)),
    [designSystemProject, projectFiles],
  );
  // Only the connect-repo CTA copy depends on this (connect vs re-import), so
  // resolve it lazily and only while the CTA is actually showing. Tri-state:
  // `undefined` means the status fetch has not resolved yet, which keeps the
  // CTA neutral and disabled so a fast click can't fire the wrong action.
  const [githubConnected, setGithubConnected] = useState<boolean | undefined>(undefined);
  useEffect(() => {
    if (!connectRepoNeeded) {
      setGithubConnected(undefined);
      return;
    }
    let aborted = false;
    const controller = new AbortController();
    const refresh = () => {
      void fetchConnectorStatuses({ signal: controller.signal }).then((statuses) => {
        if (!aborted) setGithubConnected(statuses.github?.status === 'connected');
      });
    };
    refresh();
    // Connecting GitHub happens in the Connectors dialog or an external OAuth
    // window, neither of which changes connectRepoNeeded. Re-check on focus so
    // the CTA flips from "Connect GitHub" to "Import repo" when the user returns.
    const onFocus = () => refresh();
    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onFocus);
    return () => {
      aborted = true;
      controller.abort();
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onFocus);
    };
  }, [connectRepoNeeded]);

  // Signal that pushes a draft into the chat composer (the "Import repo" CTA).
  const [composerDraftSignal, setComposerDraftSignal] = useState<{ text: string; nonce: number }>();
  // One handler for both the review banner and the chat CTA. When GitHub is
  // not connected it opens Connectors; once connected it prefills the composer
  // with the import instruction so the user can review and send it.
  const handleConnectRepo = useCallback(() => {
    // Status not resolved yet; the CTA is disabled in this window, but guard
    // anyway so a stray call can't route a connected account to Connectors.
    if (githubConnected === undefined) return;
    if (githubConnected) {
      setComposerDraftSignal({
        text: buildRepoImportPrompt(designSystemProject, projectFiles.map((file) => file.name)),
        nonce: Date.now(),
      });
    } else {
      onOpenSettings('composio');
    }
  }, [githubConnected, onOpenSettings, designSystemProject, projectFiles]);

  // "Next step" affordance handlers (shown under the last assistant message
  // once it produced a previewable HTML artifact). Share reuses the preview
  // workspace's existing Share/Export menu. The featured design-toolbox rows are
  // driven by ChatPane's composer ref, so ProjectView no longer wires them here.
  const handleArtifactShare = useCallback(
    (fileName: string) => {
      requestOpenFile(fileName);
      setShareRequest({ name: fileName, nonce: Date.now() });
    },
    [requestOpenFile],
  );
  // Mirrors share, but opens the workspace's Download/Export menu (PDF / image /
  // zip / standalone HTML / save-as-template) instead of a bare file download.
  const handleArtifactDownload = useCallback(
    (fileName: string) => {
      requestOpenFile(fileName);
      setDownloadRequest({ name: fileName, nonce: Date.now() });
    },
    [requestOpenFile],
  );

  const handleBrowserUsePrompt = useCallback((text: string) => {
    setWorkspaceFocused(false);
    setComposerDraftSignal({
      text,
      nonce: Date.now(),
    });
  }, []);

  const isDeck = useMemo(
    () =>
      (skills.find((s) => s.id === project.skillId) ??
        designTemplates.find((s) => s.id === project.skillId))?.mode === 'deck',
    [skills, designTemplates, project.skillId],
  );
  const chatResizeLabel = t('project.resizeChatPanel');
  const workspacePanelTrack =
    workspacePanelMinWidth === 0
      ? 'minmax(0, 1fr)'
      : `minmax(${workspacePanelMinWidth}px, 1fr)`;
  const splitLeftPanelWidth = leftInspectorActive
    ? COMMENT_INSPECTOR_PANEL_WIDTH
    : chatPanelWidthRef.current;
  const chatPanelAriaMinWidth = Math.min(MIN_CHAT_PANEL_WIDTH, chatPanelMaxWidth);

  const renderPreferredChatPanelWidth = useCallback((
    preferredWidth: number,
    maxWidth = chatPanelMaxWidthRef.current,
    options: { commitState?: boolean } = {},
  ): number => {
    const next = clampChatPanelWidth(preferredWidth, maxWidth);
    chatPanelWidthRef.current = next;
    applySplitChatPanelWidth(splitRef.current, next, workspacePanelTrack);
    if (options.commitState !== false) setChatPanelWidth(next);
    return next;
  }, [workspacePanelTrack]);

  const applyChatPanelWidth = useCallback((
    width: number,
    options: { commitState?: boolean } = {},
  ): number => {
    const nextPreferred = clampPreferredChatPanelWidth(
      clampChatPanelWidth(width, chatPanelMaxWidthRef.current),
    );
    preferredChatPanelWidthRef.current = nextPreferred;
    return renderPreferredChatPanelWidth(nextPreferred, chatPanelMaxWidthRef.current, options);
  }, [renderPreferredChatPanelWidth]);

  const finishChatPanelResize = useCallback((saveFinalWidth = true) => {
    pointerCleanupRef.current?.();
    pointerCleanupRef.current = null;
    if (pointerFrameRef.current !== null) {
      cancelAnimationFrame(pointerFrameRef.current);
      pointerFrameRef.current = null;
    }
    pendingPointerClientXRef.current = null;
    resizeStateRef.current = null;
    setResizingChatPanel(false);
    if (saveFinalWidth) {
      const finalWidth = renderPreferredChatPanelWidth(preferredChatPanelWidthRef.current);
      saveChatPanelWidth(finalWidth);
    }
  }, [renderPreferredChatPanelWidth]);

  useEffect(() => {
    chatPanelWidthRef.current = chatPanelWidth;
    applySplitChatPanelWidth(splitRef.current, chatPanelWidth, workspacePanelTrack);
  }, [chatPanelWidth, workspacePanelTrack]);

  useEffect(() => {
    chatPanelMaxWidthRef.current = chatPanelMaxWidth;
  }, [chatPanelMaxWidth]);

  useLayoutEffect(() => {
    const split = splitRef.current;
    if (!split) return undefined;

    const updateAllowedWidth = () => {
      const splitWidth = split.clientWidth;
      const nextWorkspaceMin = workspacePanelMinWidthForSplit(splitWidth);
      const nextMax = maxChatPanelWidthForSplit(splitWidth);
      chatPanelMaxWidthRef.current = nextMax;
      setWorkspacePanelMinWidth(nextWorkspaceMin);
      setChatPanelMaxWidth(nextMax);
      renderPreferredChatPanelWidth(preferredChatPanelWidthRef.current, nextMax);
    };

    updateAllowedWidth();

    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(updateAllowedWidth);
      observer.observe(split);
      return () => observer.disconnect();
    }

    window.addEventListener('resize', updateAllowedWidth);
    return () => window.removeEventListener('resize', updateAllowedWidth);
  }, [renderPreferredChatPanelWidth]);

  useEffect(() => () => finishChatPanelResize(false), [finishChatPanelResize]);

  const handleChatResizePointerDown = useCallback((event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const split = splitRef.current;
    if (!split) return;
    event.preventDefault();
    event.currentTarget.focus();
    event.currentTarget.setPointerCapture(event.pointerId);
    pointerCleanupRef.current?.();
    setResizingChatPanel(true);
    resizeStartPreferredWidthRef.current = preferredChatPanelWidthRef.current;

    const updateWidthFromClientX = (clientX: number) => {
      const state = resizeStateRef.current;
      if (!state) return;
      const delta = clientX - state.startClientX;
      if (delta === 0 && !state.hasMoved) return;
      state.hasMoved = true;
      const rawWidth = state.startWidth + (state.isRtl ? -delta : delta);
      applyChatPanelWidth(rawWidth, { commitState: false });
    };

    const flushPendingPointerMove = () => {
      if (pointerFrameRef.current !== null) {
        cancelAnimationFrame(pointerFrameRef.current);
        pointerFrameRef.current = null;
      }
      const clientX = pendingPointerClientXRef.current;
      pendingPointerClientXRef.current = null;
      if (clientX !== null) updateWidthFromClientX(clientX);
    };

    resizeStateRef.current = {
      startClientX: event.clientX,
      startWidth: chatPanelWidthRef.current,
      isRtl: window.getComputedStyle(split).direction === 'rtl',
      hasMoved: false,
    };

    const handlePointerMove = (moveEvent: PointerEvent) => {
      pendingPointerClientXRef.current = moveEvent.clientX;
      if (pointerFrameRef.current !== null) return;
      pointerFrameRef.current = requestAnimationFrame(() => {
        pointerFrameRef.current = null;
        flushPendingPointerMove();
      });
    };
    const handlePointerEnd = () => {
      flushPendingPointerMove();
      finishChatPanelResize(true);
    };
    const handlePointerCancel = () => {
      flushPendingPointerMove();
      preferredChatPanelWidthRef.current = resizeStartPreferredWidthRef.current;
      renderPreferredChatPanelWidth(resizeStartPreferredWidthRef.current);
      finishChatPanelResize(false);
    };
    const cleanup = () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerEnd);
      window.removeEventListener('pointercancel', handlePointerCancel);
      window.removeEventListener('blur', handlePointerCancel);
    };

    pointerCleanupRef.current = cleanup;
    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerEnd);
    window.addEventListener('pointercancel', handlePointerCancel);
    window.addEventListener('blur', handlePointerCancel);
  }, [applyChatPanelWidth, finishChatPanelResize, renderPreferredChatPanelWidth]);

  const handleChatResizeBlur = useCallback(() => {
    if (!pointerCleanupRef.current) return;
    preferredChatPanelWidthRef.current = resizeStartPreferredWidthRef.current;
    renderPreferredChatPanelWidth(resizeStartPreferredWidthRef.current);
    finishChatPanelResize(false);
  }, [finishChatPanelResize, renderPreferredChatPanelWidth]);

  const handleChatResizeKeyDown = useCallback((event: ReactKeyboardEvent<HTMLDivElement>) => {
    let nextWidth: number | null = null;
    const split = splitRef.current;
    const isRtl = split ? window.getComputedStyle(split).direction === 'rtl' : false;
    if (event.key === 'ArrowLeft') {
      nextWidth = chatPanelWidthRef.current + (isRtl ? 1 : -1) * CHAT_PANEL_KEYBOARD_STEP;
    } else if (event.key === 'ArrowRight') {
      nextWidth = chatPanelWidthRef.current + (isRtl ? -1 : 1) * CHAT_PANEL_KEYBOARD_STEP;
    } else if (event.key === 'Home') {
      nextWidth = MIN_CHAT_PANEL_WIDTH;
    } else if (event.key === 'End') {
      nextWidth = chatPanelMaxWidthRef.current;
    }
    if (nextWidth === null) return;
    event.preventDefault();
    const next = applyChatPanelWidth(nextWidth);
    saveChatPanelWidth(next);
  }, [applyChatPanelWidth]);

  // Hand the pending prompt to ChatPane exactly once per project. The local
  // project-scoped snapshot survives the conversation-id remount, while the
  // persisted pendingPrompt is cleared so refreshes and later entries do not
  // re-seed the composer.
  //
  // PluginLoopHome auto-send case: when the project was created with
  // `autoSendFirstMessage`, app.tsx left a sessionStorage flag telling us
  // to fire the prompt as a real user message immediately. We must NOT
  // seed initialDraft in that case — otherwise the textarea echoes the
  // prompt while it is also streaming as the first user message. The ref
  // captures the prompt independently so downstream effects can still
  // dispatch the auto-send without going through initialDraft.
  const autoSendSeedRef = useRef<string | null>(null);
  const autoSendAttachmentsRef = useRef<ChatAttachment[] | null>(null);
  const autoSendFirstMessageRef = useRef(false);
  if (autoSendSeedRef.current === null) {
    let isAutoSend = false;
    try {
      isAutoSend = Boolean(
        window.sessionStorage.getItem(autoSendFirstMessageKey(project.id)),
      );
    } catch {
      /* sessionStorage may be unavailable; treat as manual flow. */
    }
    autoSendFirstMessageRef.current = isAutoSend;
    autoSendSeedRef.current = isAutoSend ? (project.pendingPrompt ?? '') : '';
    autoSendAttachmentsRef.current = isAutoSend ? readAutoSendAttachments(project.id) : [];
  }
  const [initialDraft, setInitialDraft] = useState<
    { projectId: string; value: string } | undefined
  >(
    autoSendSeedRef.current || !project.pendingPrompt
      ? undefined
      : { projectId: project.id, value: project.pendingPrompt },
  );
  useEffect(() => {
    const pendingPrompt = project.pendingPrompt;
    if (!pendingPrompt) return;
    if (autoSendFirstMessageRef.current) {
      onClearPendingPrompt();
      return;
    }
    setInitialDraft((current) =>
      current?.projectId === project.id
        ? current
        : { projectId: project.id, value: pendingPrompt },
    );
    onClearPendingPrompt();
  }, [project.id, project.pendingPrompt, onClearPendingPrompt]);
  const chatInitialDraft =
    chatSeed?.value ?? (initialDraft?.projectId === project.id ? initialDraft.value : undefined);

  // Continue in CLI / Finalize design package handlers + keyboard
  // shortcut wiring. Close to the JSX so the data flow is easy to
  // trace from the toolbar back to its sources.
  const handleFinalize = useCallback(() => {
    const request = buildFinalizeRequest(config);
    if (!request) {
      setProjectActionsToast(buildFinalizeCredentialsMissingToast(config));
      return;
    }
    void finalize.trigger(request).then((result) => {
      if (result) void designMdState.refresh();
    });
  }, [finalize, config, designMdState]);

  const handleCancelFinalize = useCallback(() => {
    finalize.cancel();
  }, [finalize]);

  const handleContinueInCli = useCallback(async () => {
    const projectDir = projectDetail.resolvedDir;
    if (!projectDir) {
      setProjectActionsToast({
        message: 'Working directory unavailable. Update the daemon to enable Continue in CLI.',
        details: null,
      });
      return;
    }
    const prompt = buildClipboardPrompt({
      project: { id: project.id, name: project.name },
      designMdState: {
        generatedAt: designMdState.generatedAt,
        transcriptMessageCount: designMdState.transcriptMessageCount,
        designSystemId: designMdState.designSystemId,
        currentArtifact: designMdState.currentArtifact,
      },
      projectDir,
    });
    const copied = await copyToClipboard(prompt);
    if (!copied) {
      // Clipboard write failed in both the canonical and execCommand
      // fallback paths (locked clipboard / insecure context). Surface
      // the prompt body in the toast so the user can manually
      // select-and-copy. Do not open the folder — the user has nothing
      // to paste yet.
      setProjectActionsToast({
        message: 'Clipboard unavailable. Copy this prompt manually, then run `claude` at the working directory.',
        details: `Working directory: ${projectDir}`,
        code: prompt,
      });
      return;
    }
    const launched = await terminalLauncher.open(project.id);
    setProjectActionsToast(buildContinueInCliToast(projectDir, launched));
  }, [
    project.id,
    project.name,
    projectDetail.resolvedDir,
    designMdState.generatedAt,
    designMdState.transcriptMessageCount,
    designMdState.designSystemId,
    designMdState.currentArtifact,
    terminalLauncher,
  ]);

  // Defensive: if the conversation already has messages once they
  // hydrate, the pendingPrompt that seeded the composer is stale (the
  // user sent it earlier but onClearPendingPrompt did not get a chance
  // to patch the server before the page reloaded). Drop the seed so the
  // textarea does not echo a prompt the user already submitted.
  useEffect(() => {
    if (initialDraft && messages.length > 0) {
      setInitialDraft(undefined);
    }
  }, [initialDraft, messages.length]);

  // §8.4 — when the project was created with a plugin pinned (the
  // PluginLoopHome → POST /api/projects path), fetch the immutable
  // snapshot once so ChatPane can render the active plugin as a
  // context chip on user messages instead of re-rendering the inline
  // plugin rail. Re-fetches when the pinned id changes; cancelled if
  // the project switches away mid-flight to avoid setState-on-unmount.
  const [activePluginSnapshot, setActivePluginSnapshot] =
    useState<AppliedPluginSnapshot | null>(null);
  const [contextPluginDetails, setContextPluginDetails] =
    useState<InstalledPluginRecord | null>(null);
  const [contextDesignSystemDetails, setContextDesignSystemDetails] =
    useState<DesignSystemSummary | null>(null);
  useEffect(() => {
    const snapshotId = project.appliedPluginSnapshotId;
    if (!snapshotId) {
      setActivePluginSnapshot(null);
      return;
    }
    let cancelled = false;
    void fetchAppliedPluginSnapshot(snapshotId).then((snap) => {
      if (cancelled) return;
      setActivePluginSnapshot(snap);
    });
    return () => {
      cancelled = true;
    };
  }, [project.appliedPluginSnapshotId]);
  const handleOpenContextPluginDetails = useCallback(async (pluginId: string) => {
    const normalizedId = pluginId.trim();
    if (!normalizedId) return;
    const plugins = await listPlugins({ includeHidden: true });
    const record = plugins.find((plugin) => plugin.id === normalizedId);
    if (record) setContextPluginDetails(record);
  }, []);
  const chatDesignSystemSummary = useMemo(() => {
    if (activeDesignSystemSummary) return activeDesignSystemSummary;
    const designSystemName = activePluginSnapshot?.inputs?.designSystem;
    if (typeof designSystemName !== 'string') return null;
    const normalized = designSystemName.trim();
    if (!normalized || normalized === 'the active project design system') return null;
    return designSystems.find((d) => d.title === normalized) ?? null;
  }, [activeDesignSystemSummary, activePluginSnapshot?.inputs, designSystems]);

  // Lift finalize errors into the shared project-actions toast so the
  // user sees both the daemon's category message and any upstream
  // detail (per #450 verification commitment).
  useEffect(() => {
    if (finalize.error) {
      setProjectActionsToast({
        message: finalize.error.message,
        details: finalize.error.details,
      });
    }
  }, [finalize.error]);

  // ⌘+Shift+K (mac) / Ctrl+Shift+K (others) → Continue in CLI. Mirrors
  // the capture-phase, platform-gated pattern from FileWorkspace's
  // Quick Switcher shortcut. ⌘+Shift+K is free (⌘+P is the only
  // existing primary-modifier shortcut on this surface).
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const primary = isMacPlatform() ? e.metaKey && !e.ctrlKey : e.ctrlKey && !e.metaKey;
      if (primary && e.shiftKey && !e.altKey && e.key.toLowerCase() === 'k') {
        if (e.isComposing) return;
        if (!designMdState.exists) return;
        e.preventDefault();
        void handleContinueInCli();
      }
    };
    window.addEventListener('keydown', onKeyDown, { capture: true });
    return () => window.removeEventListener('keydown', onKeyDown, { capture: true });
  }, [designMdState.exists, handleContinueInCli]);

  // PluginLoopHome auto-send: when the user submits on Home, app.tsx
  // sets `sessionStorage['od:auto-send-first:<projectId>']` and routes
  // through createProject. Once the conversation id resolves and the
  // composer is mounted, fire handleSend(pendingPrompt) exactly once so
  // the user lands inside a running pipeline without an extra click.
  // We gate on `messages.length === 0` so a refresh after the run is
  // mid-flight never double-fires; the sessionStorage flag is cleared
  // immediately after the first dispatch.
  const autoSentRef = useRef(false);
  useEffect(() => {
    if (autoSentRef.current) return;
    if (!activeConversationId) return;
    // Wait for the initial listMessages DB read to land. Without this gate
    // the auto-send fires before the in-flight DB response, which then
    // arrives with `setMessages([])` and wipes the freshly-pushed user +
    // assistant placeholder out of React state — leaving the daemon's run
    // with no in-memory message to attach the runId to.
    if (!messagesInitialized) return;
    if (streaming) return;
    if (messages.length > 0) return;
    let flag: string | null = null;
    try {
      flag = window.sessionStorage.getItem(autoSendFirstMessageKey(project.id));
    } catch {
      flag = null;
    }
    if (!flag) return;
    // Prefer the seed captured at mount (autoSendSeedRef) — it survives
    // even after onClearPendingPrompt wipes project.pendingPrompt on the
    // server. Fall back to the live values for any edge case where the
    // ref was not populated (e.g. sessionStorage error path).
    const seed = (
      autoSendSeedRef.current ||
      (initialDraft?.projectId === project.id ? initialDraft.value : '') ||
      project.pendingPrompt ||
      ''
    ).trim();
    const attachments = autoSendAttachmentsRef.current ?? [];
    if (!seed && attachments.length === 0) {
      autoSentRef.current = true;
      clearAutoSendSession(project.id);
      return;
    }
    autoSentRef.current = true;
    if (isDesignSystemWorkspaceMetadata(project.metadata)) {
      markDesignSystemAuditAutoRepairEligible(project.id);
    }
    clearAutoSendSession(project.id);
    autoSendAttachmentsRef.current = [];
    void handleSend(seed, attachments, []);
  }, [
    activeConversationId,
    messagesInitialized,
    streaming,
    messages.length,
    project.id,
    project.metadata,
    initialDraft,
    project.pendingPrompt,
    handleSend,
  ]);

  // Wire the Critique Theater drop-in mount into the project workspace.
  // The hook reads the M1 Settings toggle out of the existing
  // `open-design:config` localStorage blob and stays in sync with the
  // platform `storage` event (cross-tab) plus the same-tab
  // `open-design:critique-theater-toggle` CustomEvent. The mount itself
  // returns `null` until the daemon emits a `critique.run_started` for
  // the active project, so the visual surface is unchanged for users
  // who have not opted in. The daemon-side gate
  // (`isCritiqueEnabled(...)` in `apps/daemon/src/server.ts`) is the
  // authority for whether a run is actually wired through the critique
  // pipeline; this hook only governs whether the web layer renders the
  // resulting SSE stream.
  const critiqueTheaterEnabled = useCritiqueTheaterEnabled();

  // CLI / agent selector lives below the chat conversation (composer footer),
  // not in the top-right header.
  const executionControls = (
    <AvatarMenu
      config={config}
      agents={agents}
      daemonLive={daemonLive}
      onModeChange={onModeChange}
      onOpen={() => {
        trackComposerBarClick(analytics.track, {
          page_name: 'chat_panel',
          area: 'chat_composer',
          element: 'agent_selector_open',
          ...(project?.id ? { project_id: project.id } : {}),
        });
      }}
      onAgentChange={(id) => {
        trackComposerBarClick(analytics.track, {
          page_name: 'chat_panel',
          area: 'chat_composer',
          element: 'agent_select',
          agent_id: id,
          ...(project?.id ? { project_id: project.id } : {}),
        });
        onAgentChange(id);
      }}
      onAgentModelChange={(agentId, choice) => {
        trackComposerBarClick(analytics.track, {
          page_name: 'chat_panel',
          area: 'chat_composer',
          element: 'agent_model_select',
          agent_id: agentId,
          ...(choice?.model ? { model_id: choice.model } : {}),
          ...(project?.id ? { project_id: project.id } : {}),
        });
        onAgentModelChange(agentId, choice);
      }}
      onApiModelChange={(model) => {
        trackComposerBarClick(analytics.track, {
          page_name: 'chat_panel',
          area: 'chat_composer',
          element: 'agent_model_select',
          model_id: model,
          ...(project?.id ? { project_id: project.id } : {}),
        });
        onApiModelChange?.(model);
      }}
      onOpenSettings={onOpenSettings}
      onRefreshAgents={onRefreshAgents}
      placement="up"
    />
  );

  return (
    <div className="app">
      <CritiqueTheaterMount
        projectId={project.id}
        enabled={critiqueTheaterEnabled}
      />
      {/* ProjectActionsToolbar removed per 00efdcba — hide finalize-design
          toolbar from project header. Restore from cf1cd9bb if product
          wants the Finalize + Continue-in-CLI buttons back in the chrome. */}
      <div
        ref={splitRef}
        className={[
          projectSplitClassName(workspaceFocused),
          leftInspectorActive && !workspaceFocused ? 'split-manual-edit' : '',
          resizingChatPanel && !workspaceFocused ? 'is-resizing-chat' : '',
        ].filter(Boolean).join(' ')}
        style={projectSplitStyle(workspaceFocused, splitLeftPanelWidth, workspacePanelTrack)}
      >
        <div className="split-chat-slot" hidden={workspaceFocused}>
          {commentInspectorActive ? (
            <div
              id={commentInspectorPortalId}
              className="comment-left-host"
              aria-label="Comments"
            />
          ) : activeConversationId || conversationLoadError ? (
            <ChatPane
              // The conversation id is part of the key so switching conversations
              // resets internal scroll/draft state inside ChatPane and ChatComposer.
              key={`${project.id}:${activeConversationId ?? 'conversation-unavailable'}:${chatSeed?.id ?? 'ready'}`}
              messages={messages}
              streaming={currentConversationStreaming}
              liveToolInput={liveToolInput}
              loading={currentConversationLoading}
              sendDisabled={currentConversationSendDisabled}
              queuedItems={currentConversationQueuedItems}
              error={conversationLoadError ?? error ?? audioVoiceOptionsError}
              projectId={project.id}
              sessionMode={activeSessionMode}
              onSessionModeChange={handleActiveConversationSessionModeChange}
              projectKindForTracking={projectKindToTracking(project.metadata?.kind, project.metadata?.videoModel)}
              projectFiles={projectFiles}
              activeProjectFileName={activeProjectFileName}
              hasActiveDesignSystem={!!project.designSystemId}
              activeDesignSystem={chatDesignSystemSummary}
              projectFileNames={projectFileNames}
              skills={skills}
              onEnsureProject={handleEnsureProject}
              previewComments={previewComments}
              attachedComments={attachedComments}
              onAttachComment={attachPreviewComment}
              onDetachComment={detachPreviewComment}
              onDeleteComment={(commentId) => void removePreviewComment(commentId)}
              onSend={handleSend}
              onRetry={handleRetry}
              onResumeRun={handleResumeRun}
              onStop={handleStop}
              onRemoveQueuedSend={removeQueuedChatSend}
              onUpdateQueuedSend={updateQueuedChatSend}
              onReorderQueuedSends={reorderCurrentConversationQueuedChatSends}
              onSendQueuedNow={sendQueuedChatSendNow}
              onRequestOpenFile={requestOpenFile}
              onRequestPluginDetails={handleOpenContextPluginDetails}
              onRequestDesignSystemDetails={setContextDesignSystemDetails}
              onRequestPluginFolderAgentAction={handlePluginFolderAgentAction}
              activePluginActionPaths={activePluginActionPaths}
              hiddenPluginActionPaths={hiddenAssistantPluginActionPaths}
              onShareToOpenDesign={handleShareToOpenDesign}
              shareToOpenDesignBusyMessageId={shareToOpenDesignBusyMessageId}
              forceStreamingMessageIds={forceStreamingPluginMessageIds}
              initialDraft={chatInitialDraft}
              onOpenQuestions={openQuestionsTab}
              onContinueRemainingTasks={handleContinueRemainingTasks}
              onAssistantFeedback={handleAssistantFeedback}
              onArtifactShare={handleArtifactShare}
              onArtifactDownload={handleArtifactDownload}
              onForkFromMessage={handleForkFromMessage}
              forkingMessageId={forkingMessageId}
              onNewConversation={handleNewConversation}
              newConversationDisabled={newConversationDisabled}
              conversations={conversations}
              activeConversationId={activeConversationId}
              messagesConversationId={messagesConversationId}
              onSelectConversation={handleSelectConversation}
              onDeleteConversation={handleDeleteConversation}
              config={config}
              onOpenSettings={onOpenSettings}
              showByokRecoveryAction={
                config.mode === 'api' &&
                daemonLive &&
                (
                  !config.apiKey.trim() ||
                  !config.baseUrl.trim() ||
                  !config.model.trim()
                )
              }
              onSwitchToLocalCli={() => {
                setError(null);
                onModeChange('daemon');
              }}
              onOpenAmrSettings={onOpenAmrSettings}
              onSwitchToAmrAndRetry={handleSwitchToAmrAndRetry}
              onLaunchAntigravityOauth={handleLaunchAntigravityOauth}
              onOpenMcpSettings={onOpenMcpSettings}
              onBrowsePlugins={onBrowsePlugins}
              onOpenConnectors={onOpenConnectors}
              connectRepoNeeded={connectRepoNeeded}
              githubConnected={githubConnected}
              onConnectRepo={handleConnectRepo}
              composerDraftSignal={composerDraftSignal}
              petConfig={config.pet}
              onAdoptPet={onAdoptPetInline}
              onTogglePet={onTogglePet}
              onOpenPetSettings={onOpenPetSettings}
              researchAvailable={config.mode === 'daemon'}
              byokApiProtocol={config.apiProtocol}
              byokImageModel={byokImageModelOverride}
              onChangeByokImageModel={setByokImageModelOverride}
              byokVideoModel={byokVideoModelOverride}
              onChangeByokVideoModel={setByokVideoModelOverride}
              byokSpeechModel={byokSpeechModelOverride}
              onChangeByokSpeechModel={setByokSpeechModelOverride}
              byokSpeechVoice={byokSpeechVoiceOverride}
              onChangeByokSpeechVoice={setByokSpeechVoiceOverride}
              projectMetadata={project.metadata}
              onProjectMetadataChange={(metadata) => {
                onProjectChange({ ...project, metadata });
              }}
              activeWorkspaceContext={activeWorkspaceContext}
              workspaceContexts={workspaceContexts}
              currentSkillId={project.skillId}
              onProjectSkillChange={(skillId) => {
                onProjectChange({ ...project, skillId });
              }}
              activePluginSnapshot={activePluginSnapshot}
              currentDesignSystemId={project.designSystemId}
              onActiveDesignSystemChange={(updatedProject) => {
                onProjectChange(updatedProject);
              }}
              onShowToast={(message) => {
                setProjectActionsToast({ message, details: null });
              }}
              onBack={onBack}
              backLabel={t('project.backToProjects')}
              composerFooterAccessory={executionControls}
              projectHeader={(
                <span className="chat-project-title-line">
                  <span
                    className="title editable"
                    data-testid="project-title"
                    title={project.name}
                    tabIndex={0}
                    role="textbox"
                    suppressContentEditableWarning
                    contentEditable
                    onBlur={(e) => handleProjectRename(e.currentTarget.textContent ?? '')}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        (e.currentTarget as HTMLElement).blur();
                      }
                    }}
                  >
                    {project.name}
                  </span>
                  {projectMeta !== t('project.metaFreeform') ? (
                    <span className="meta" data-testid="project-meta">{projectMeta}</span>
                  ) : null}
                </span>
              )}
              designSystemPicker={(
                <DesignSystemPicker
                  designSystems={designSystems}
                  selectedId={project.designSystemId ?? null}
                  onChange={handleChangeDesignSystemId}
                />
              )}
            />
          ) : (
            <div className="pane" data-testid="chat-pane-loading">
              <CenteredLoader />
            </div>
          )}
        </div>
        {!workspaceFocused ? (
          leftInspectorActive ? (
            <div className="split-edit-divider" aria-hidden />
          ) : (
            <div
              className="split-resize-handle"
              role="separator"
              aria-orientation="vertical"
              aria-label={chatResizeLabel}
              aria-valuemin={chatPanelAriaMinWidth}
              aria-valuemax={chatPanelMaxWidth}
              aria-valuenow={chatPanelWidth}
              tabIndex={0}
              title={chatResizeLabel}
              onPointerDown={handleChatResizePointerDown}
              onKeyDown={handleChatResizeKeyDown}
              onBlur={handleChatResizeBlur}
            />
          )
        ) : null}
        <FileWorkspace
          projectId={project.id}
          projectKind={projectKindToTracking(project.metadata?.kind, project.metadata?.videoModel) ?? 'prototype'}
          rootDirName={(() => {
            const baseDir =
              projectDetail.project?.metadata?.baseDir ?? project.metadata?.baseDir;
            return typeof baseDir === 'string'
              ? baseDir.split(/[/\\]/).filter(Boolean).pop()
              : undefined;
          })()}
          reloading={false}
          resolvedDir={projectDetail.resolvedDir}
          files={projectFiles}
          liveArtifacts={liveArtifacts}
          filesRefreshKey={filesRefresh}
          onRefreshFiles={() => {
            void refreshWorkspaceItems();
          }}
          isDeck={isDeck}
          onExportAsPptx={handleExportAsPptx}
          streaming={currentConversationActionDisabled}
          commentQueueOnSend={commentQueueOnSend}
          commentSendDisabled={currentConversationQueueDisabled}
          openRequest={openRequest}
          shareRequest={shareRequest}
          downloadRequest={downloadRequest}
          slideNavRequest={slideNavRequest}
          liveArtifactEvents={liveArtifactEvents}
          designSystemActivityEvents={designSystemActivityEvents}
          tabsState={openTabsState}
          onTabsStateChange={persistTabsState}
          previewComments={previewComments}
          onSavePreviewComment={savePreviewComment}
          onRemovePreviewComment={removePreviewComment}
          onSendBoardCommentAttachments={handleSendBoardCommentAttachments}
          onRequestBrowserUsePrompt={handleBrowserUsePrompt}
          onPluginFolderAgentAction={handlePluginFolderAgentAction}
          activePluginActionPaths={activePluginActionPaths}
          preferredPreviewFile={project.metadata?.entryFile ?? null}
          autoPreviewDesignArtifacts={project.metadata?.importedFrom === 'folder'}
          focusMode={workspaceFocused}
          onFocusModeChange={setWorkspaceFocused}
          designSystemProject={designSystemProject}
          defaultDesignSystemId={config.designSystemId}
          onSetDefaultDesignSystem={onChangeDefaultDesignSystem}
          onDesignSystemsRefresh={onDesignSystemsRefresh}
          onDesignSystemNeedsWork={sendDesignSystemFeedback}
          designSystemReview={project.metadata?.designSystemReview}
          onDesignSystemReviewDecision={persistDesignSystemReviewDecision}
          onConnectRepo={handleConnectRepo}
          githubConnected={githubConnected}
          commentPortalId={commentInspectorPortalId}
          onCommentModeChange={setCommentInspectorActive}
          chatConfig={config}
          chatAgentsById={agentsById}
          chatLocale={locale}
          conversations={conversations}
          activeConversationId={activeConversationId}
          onSelectConversation={handleSelectConversation}
          onDeleteConversation={handleDeleteConversation}
          onRenameConversation={handleRenameConversation}
          onConversationSessionModeChange={handleConversationSessionModeChange}
          onNewConversation={handleNewConversation}
          activeConversationChat={activeConversationChatState}
          onActiveContextChange={handleActiveWorkspaceContextChange}
          onWorkspaceContextsChange={handleWorkspaceContextsChange}
          messages={messages}
          artifactHtml={artifact?.html}
          conversationError={error}
          onRetry={handleRetry}
          onAuthorizeAndRetry={handleSwitchToAmrAndRetry}
          onLaunchTerminalAuth={handleLaunchAntigravityOauth}
          conversationId={activeConversationId}
          headerActions={(
            <>
              <HandoffButton
                projectId={project.id}
                projectName={project.name}
                projectDir={projectDetail.resolvedDir}
                agents={agents}
                artifactId={headerArtifact.artifact_id}
                artifactKind={headerArtifact.artifact_kind}
                metricsConsent={config.telemetry?.metrics === true}
                installationId={config.installationId}
              />
              <EntrySettingsMenu
                config={config}
                onThemeChange={handleThemeChange}
                onOpenSettings={onOpenSettings}
                trackingPageName="artifact"
                onTrackTriggerClick={() => {
                  // Spec row 52: the settings gear in the artifact header.
                  // Carry the active artifact so settings slices line up with
                  // the rest of the artifact_header funnel.
                  trackArtifactHeaderClick(analytics.track, {
                    page_name: 'artifact',
                    area: 'artifact_header',
                    element: 'settings',
                    ...headerArtifact,
                  });
                }}
              />
            </>
          )}
          questionForm={displayedQuestionForm}
          questionFormPreview={displayedQuestionFormPreview}
          questionFormKey={displayedQuestionFormKey}
          questionFormInteractive={displayedQuestionFormActive}
          questionFormSubmitDisabled={currentConversationActionDisabled}
          questionFormSubmittedAnswers={displayedQuestionFormSubmittedAnswers}
          questionsGenerating={displayedQuestionsGenerating}
          focusQuestionsRequest={focusQuestionsRequest}
          onSubmitQuestionForm={(text) => {
            if (currentConversationActionDisabled) return;
            // Submitting question-form answers is a clarification turn, not a
            // fresh create/edit — tag entry_from so the dashboard can separate it.
            void handleSend(text, [], [], { entryFrom: 'question_answer' });
          }}
        />
      </div>
      {contextPluginDetails ? (
        <PluginDetailsModal
          record={contextPluginDetails}
          onClose={() => setContextPluginDetails(null)}
          onUse={() => setContextPluginDetails(null)}
          isApplying={false}
          hideUseAction
        />
      ) : null}
      {contextDesignSystemDetails ? (
        <DesignSystemPreviewModal
          system={contextDesignSystemDetails}
          onClose={() => setContextDesignSystemDetails(null)}
        />
      ) : null}
      <AnimatePresence>
        {projectActionsToast ? (
          <Toast
            message={projectActionsToast.message}
            details={projectActionsToast.details}
            code={projectActionsToast.code}
            onDismiss={() => setProjectActionsToast(null)}
          />
        ) : null}
      </AnimatePresence>
    </div>
  );
}

function artifactExtensionFor(art: Artifact): '.html' | '.jsx' | '.tsx' {
  const type = (art.artifactType || '').toLowerCase();
  const identifier = (art.identifier || '').toLowerCase();
  if (type.includes('tsx') || identifier.endsWith('.tsx')) return '.tsx';
  if (type.includes('jsx') || type.includes('react') || identifier.endsWith('.jsx')) {
    return '.jsx';
  }
  return '.html';
}

function artifactBaseNameFor(art: Artifact): string {
  return (
    (art.identifier || art.title || 'artifact')
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60) || 'artifact'
  );
}

export function findExistingArtifactProjectFile(
  art: Artifact,
  projectFiles: ProjectFile[],
  options: { minMtime?: number } = {},
): ProjectFile | null {
  const ext = artifactExtensionFor(art);
  const baseName = artifactBaseNameFor(art);
  const candidateFileName = `${baseName}${ext}`;
  const currentRunFiles = filterProjectFilesByMinMtime(projectFiles, options.minMtime);

  if (ext === '.html') {
    const pointerTarget = resolveHtmlPointerArtifactTarget({
      content: art.html,
      candidateFileName,
      projectFiles: currentRunFiles,
    });
    const pointerFile = pointerTarget
      ? currentRunFiles.find((file) => file.name === pointerTarget || file.path === pointerTarget)
      : null;
    if (pointerFile) return pointerFile;
  }

  const identifier = art.identifier || '';
  if (identifier) {
    const manifestMatches = currentRunFiles
      .filter((file) => file.artifactManifest?.metadata?.identifier === identifier)
      .sort((a, b) => b.mtime - a.mtime);
    if (manifestMatches[0]) return manifestMatches[0];
  }

  return currentRunFiles.find((file) => file.name === candidateFileName) ?? null;
}

function filterProjectFilesByMinMtime(
  projectFiles: readonly ProjectFile[],
  minMtime?: number,
): ProjectFile[] {
  return typeof minMtime === 'number' && Number.isFinite(minMtime)
    ? projectFiles.filter((file) => file.mtime >= minMtime)
    : [...projectFiles];
}

export function selectPrimaryProjectFile(files: ProjectFile[]): ProjectFile | null {
  const candidates = files
    .filter((file) => !isProcessArtifactFile(file.name))
    .map((file) => ({ file, rank: primaryProjectFileRank(file) }))
    .filter((candidate) => Number.isFinite(candidate.rank));
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => a.rank - b.rank || b.file.mtime - a.file.mtime);
  return candidates[0]?.file ?? null;
}

function isProcessArtifactFile(name: string): boolean {
  const base = name.split('/').pop()?.toLowerCase() ?? name.toLowerCase();
  return (
    base === 'critique.json'
    || base.endsWith('.log')
    || base.endsWith('.meta.json')
    || base.endsWith('.artifact.json')
    || base.endsWith('.map')
  );
}

function primaryProjectFileRank(file: ProjectFile): number {
  if (manifestDeclaresPrimary(file)) return 0;
  if (file.artifactManifest && file.artifactManifest.metadata?.inferred !== true) return 1;
  if (file.kind === 'html') return 2;
  if (file.kind === 'image') return 3;
  if (file.kind === 'video') return 4;
  if (file.kind === 'sketch') return 5;
  if (file.kind === 'pdf') return 6;
  if (file.kind === 'presentation') return 7;
  if (file.kind === 'document') return 8;
  if (file.kind === 'spreadsheet') return 9;
  return Number.POSITIVE_INFINITY;
}

function manifestDeclaresPrimary(file: ProjectFile): boolean {
  const manifest = file.artifactManifest;
  if (!manifest) return false;
  if (primaryValueTargetsFile(manifest.primary, file.name)) return true;
  const metadata = manifest.metadata;
  if (!metadata || typeof metadata !== 'object') return false;
  if (primaryValueTargetsFile(metadata.primary, file.name)) return true;
  const outputs = metadata.outputs;
  if (outputs && typeof outputs === 'object' && !Array.isArray(outputs)) {
    return primaryValueTargetsFile(
      (outputs as { primary?: unknown }).primary,
      file.name,
    );
  }
  return false;
}

function primaryValueTargetsFile(value: unknown, fileName: string): boolean {
  if (value === true) return true;
  if (typeof value !== 'string') return false;
  return normalizeProjectFileName(value) === normalizeProjectFileName(fileName);
}

function normalizeProjectFileName(value: string): string {
  return value.replace(/\\/g, '/').replace(/^\.?\//, '').toLowerCase();
}

function assistantAgentDisplayName(
  agentId: string | null,
  fallbackName?: string,
): string | undefined {
  return agentDisplayName(agentId, fallbackName) ?? undefined;
}

function isTerminalRunStatus(status: ChatMessage['runStatus']): boolean {
  return status === 'succeeded' || status === 'failed' || status === 'canceled';
}

function isActiveRunStatus(status: ChatMessage['runStatus']): boolean {
  return status === 'queued' || status === 'running';
}

export function hasRecoverableArtifactMessage(message: ChatMessage): boolean {
  if (message.role !== 'assistant') return false;
  if (!message.runId) return false;
  if (!isTerminalRunStatus(message.runStatus)) return false;
  if (message.producedFiles?.length) return false;
  const sourceText = message.content.trim().length > 0
    ? message.content
    : textContentFromAgentEvents(message.events);
  return artifactFromRecoverableSourceText(sourceText) !== null;
}

function artifactFromRecoverableSourceText(sourceText: string): Artifact | null {
  const parser = createArtifactParser();
  let parsedArtifact: Artifact | null = null;
  let liveHtml = '';
  for (const ev of [...parser.feed(sourceText), ...parser.flush()]) {
    if (ev.type === 'artifact:start') {
      liveHtml = '';
      parsedArtifact = {
        identifier: ev.identifier,
        artifactType: ev.artifactType,
        title: ev.title,
        html: '',
      };
    } else if (ev.type === 'artifact:chunk') {
      liveHtml += ev.delta;
      parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, liveHtml);
    } else if (ev.type === 'artifact:end') {
      parsedArtifact = artifactWithHtml(parsedArtifact, ev.identifier, ev.fullContent);
    }
  }
  if (parsedArtifact?.html) return parsedArtifact;

  const html = recoverStandaloneHtmlDocument(sourceText)
    ?? recoverHtmlDocumentFromMarkdownFence(sourceText);
  if (!html) return null;
  return {
    identifier: 'response',
    artifactType: 'text/html',
    title: 'Response',
    html,
  };
}

function shouldReplayTerminalRunMessage(message: ChatMessage): boolean {
  if (message.role !== 'assistant') return false;
  if (!message.runId) return false;
  if (message.runStatus !== 'succeeded') return false;
  if (message.content.trim().length > 0) return false;
  if (
    message.startedAt == null
    && !message.preTurnFileNames?.length
    && textContentFromAgentEvents(message.events).trim().length === 0
  ) {
    return false;
  }
  return !(message.producedFiles?.length);
}

function textContentFromAgentEvents(events?: AgentEvent[]): string {
  return (events ?? [])
    .filter((event): event is Extract<AgentEvent, { kind: 'text' }> => event.kind === 'text')
    .map((event) => event.text)
    .join('');
}

const QUEUED_CHAT_SENDS_STORAGE_VERSION = 1;

function queuedChatSendsStorageKey(projectId: string): string {
  return `od:chat-queued-sends:${projectId}:v${QUEUED_CHAT_SENDS_STORAGE_VERSION}`;
}

function loadQueuedChatSends(projectId: string): QueuedChatSend[] {
  if (typeof window === 'undefined') return [];
  try {
    const raw = window.localStorage.getItem(queuedChatSendsStorageKey(projectId));
    const parsed = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(isQueuedChatSend).slice(0, 100);
  } catch {
    return [];
  }
}

function saveQueuedChatSends(projectId: string, items: QueuedChatSend[]): void {
  if (typeof window === 'undefined') return;
  try {
    const key = queuedChatSendsStorageKey(projectId);
    if (items.length === 0) {
      window.localStorage.removeItem(key);
      return;
    }
    window.localStorage.setItem(key, JSON.stringify(items.slice(0, 100)));
  } catch {
    // Ignore private-mode/quota failures. The in-memory queue still works.
  }
}

function isQueuedChatSend(value: unknown): value is QueuedChatSend {
  if (typeof value !== 'object' || value == null || Array.isArray(value)) return false;
  const record = value as Partial<QueuedChatSend>;
  return (
    typeof record.id === 'string' &&
    typeof record.conversationId === 'string' &&
    typeof record.prompt === 'string' &&
    Array.isArray(record.attachments) &&
    Array.isArray(record.commentAttachments) &&
    typeof record.createdAt === 'number'
  );
}

function stripQueueOnlyFromMeta(meta: ChatSendMeta | undefined): ProjectChatSendMeta | undefined {
  if (!meta) return undefined;
  const { queueOnly: _queueOnly, ...rest } = meta;
  return Object.keys(rest).length > 0 ? rest : undefined;
}

export interface RetryTarget {
  failedAssistant: ChatMessage;
  userMsg: ChatMessage;
  priorMessages: ChatMessage[];
  preservedAttempts: ChatMessage[];
}

export function resolveRetryTarget(
  messages: ChatMessage[],
  failedAssistantId: string,
): RetryTarget | null {
  const failedIndex = messages.findIndex(
    (message) =>
      message.id === failedAssistantId &&
      message.role === 'assistant' &&
      message.runStatus === 'failed',
  );
  if (failedIndex <= 0 || failedIndex !== messages.length - 1) return null;

  let userIndex = failedIndex - 1;
  while (
    userIndex >= 0 &&
    messages[userIndex]?.role === 'assistant' &&
    messages[userIndex]?.runStatus === 'failed'
  ) {
    userIndex -= 1;
  }

  const userMsg = messages[userIndex];
  const failedAssistant = messages[failedIndex];
  if (!userMsg || userMsg.role !== 'user' || !failedAssistant) return null;

  return {
    failedAssistant,
    userMsg,
    priorMessages: messages.slice(0, userIndex),
    preservedAttempts: messages.slice(userIndex + 1, failedIndex + 1),
  };
}

function latestDesignSystemActivityEvents(messages: ChatMessage[]): AgentEvent[] {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (!message || message.role !== 'assistant') continue;
    if ((message.events?.length ?? 0) > 0) return message.events ?? [];
    if (isActiveRunStatus(message.runStatus)) return [];
  }
  return [];
}

function pluginWorkflowTitle(action: PluginFolderAgentAction): string {
  return action === 'publish' ? 'Publish repo' : 'Open Design PR';
}

function pluginWorkflowCliCommand(action: PluginFolderAgentAction, relativePath: string): string {
  return action === 'publish'
    ? `od plugin publish-repo ${relativePath}`
    : `od plugin open-design-pr ${relativePath}`;
}

function pluginWorkflowPlannedSteps(action: PluginFolderAgentAction): string[] {
  if (action === 'publish') {
    return [
      'Resolve GitHub owner and validate plugin metadata',
      'Create or update the GitHub repository',
      'Push plugin files and tags',
      'Return the repository URL',
    ];
  }
  return [
    'Ensure the Open Design fork exists',
    'Clone the fork and prepare a branch',
    'Copy the plugin into plugins/community',
    'Push the branch and open the PR form',
  ];
}

function pluginWorkflowPlannedEvents(action: PluginFolderAgentAction, relativePath: string): AgentEvent[] {
  return [
    { kind: 'text', text: `${pluginWorkflowStartContent(action, relativePath)}\n\n` },
    { kind: 'status', label: 'working', detail: pluginWorkflowTitle(action) },
  ];
}

function pluginWorkflowResultEvents(
  action: PluginFolderAgentAction,
  relativePath: string,
  message: string,
  url: string | undefined,
  log: string[] | undefined,
  ok: boolean,
  existingEvents?: AgentEvent[],
): AgentEvent[] {
  const summary = ok
    ? pluginWorkflowSuccessContent(action, relativePath, message, url, log)
    : pluginWorkflowFailureContent(action, relativePath, message, log);
  const baseEvents = (existingEvents ?? []).filter(
    (event) => !(event.kind === 'status' && event.label === 'working'),
  );
  return [
    ...baseEvents,
    { kind: 'text', text: `${summary}\n\n` },
    {
      kind: 'status',
      label: ok ? 'done' : 'failed',
      detail: ok ? 'CLI command finished' : 'CLI command failed',
    },
  ];
}

function pluginWorkflowStartContent(action: PluginFolderAgentAction, relativePath: string): string {
  const title = pluginWorkflowTitle(action);
  const command = pluginWorkflowCliCommand(action, relativePath);
  const steps = pluginWorkflowPlannedSteps(action).map((step) => `- ${step}`).join('\n');
  return `${title} started.\n\n\`\`\`bash\n${command}\n\`\`\`\n\nPlanned steps:\n${steps}`;
}

function pluginWorkflowSuccessContent(
  action: PluginFolderAgentAction,
  relativePath: string,
  message: string,
  url?: string,
  log?: string[],
): string {
  const summary = stripTrailingUrl(message, url) || `${pluginWorkflowTitle(action)} completed for \`${relativePath}\`.`;
  const lines = (log ?? []).map((line) => line.trim()).filter(Boolean).slice(0, 5);
  const command = pluginWorkflowCliCommand(action, relativePath);
  const details = lines.length > 0
    ? `\n\nCLI output:\n${lines.map((line) => `- \`${truncatePluginWorkflowLine(line)}\``).join('\n')}`
    : '';
  const link = url ? `\n\nLink: [${url}](${url})` : '';
  return `${summary}\n\n\`\`\`bash\n${command}\n\`\`\`${link}${details}`;
}

function pluginWorkflowFailureContent(
  action: PluginFolderAgentAction,
  relativePath: string,
  message: string,
  log?: string[],
): string {
  const lines = (log ?? []).map((line) => line.trim()).filter(Boolean).slice(0, 5);
  const command = pluginWorkflowCliCommand(action, relativePath);
  const details = lines.length > 0
    ? `\n\nCLI output:\n${lines.map((line) => `- \`${truncatePluginWorkflowLine(line)}\``).join('\n')}`
    : '';
  return `${pluginWorkflowTitle(action)} failed.\n\n\`\`\`bash\n${command}\n\`\`\`\n\n${message}${details}`;
}

function truncatePluginWorkflowLine(line: string): string {
  return line.length > 160 ? `${line.slice(0, 157)}...` : line;
}

function stripTrailingUrl(message: string, url?: string): string {
  const text = message.trim();
  const link = url?.trim();
  if (!link) return text;
  return text.replace(new RegExp(`\\s*${escapeRegExp(link)}\\s*$`), '').trim();
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// A daemon assistant message that is "queued/running" but has no runId yet
// is in-flight on the client: POST /api/runs has not returned. Persisting it
// in this state creates a phantom DB row that the reattach loop can never
// recover (the daemon either never saw the request or the response was lost),
// which is what produced the "Working 24m+" stuck UI. Treat the in-flight
// window as ephemeral and only write to DB once a runId pins the row to a
// real daemon run — or once the run reaches a terminal state.
function isPhantomDaemonRunMessage(m: ChatMessage): boolean {
  return (
    m.role === 'assistant' &&
    isActiveRunStatus(m.runStatus) &&
    !m.runId
  );
}

function isStoppableAssistantMessage(message: ChatMessage): boolean {
  if (message.role !== 'assistant') return false;
  if (isActiveRunStatus(message.runStatus)) return true;
  return message.runStatus === undefined && message.endedAt === undefined && message.startedAt !== undefined;
}

export function resolveSucceededRunStatus(status: ChatMessage['runStatus']): ChatMessage['runStatus'] {
  return status === 'failed' || status === 'canceled' ? status : 'succeeded';
}

export function computeProducedFiles(
  beforeNames: ReadonlySet<string> | readonly string[] | undefined,
  next: readonly ProjectFile[],
): ProjectFile[] | undefined {
  if (!beforeNames) return undefined;
  const set = beforeNames instanceof Set ? beforeNames : new Set(beforeNames);
  return filterImplicitProducedFiles(next.filter((f) => !set.has(f.name)));
}

// Reattach with a recovered (on-disk) artifact must still include any
// other files the turn produced before the artifact write — replacing
// the diff with a single file was the regression noted on PR #2383.
export function mergeRecoveredArtifact(
  diff: readonly ProjectFile[],
  recovered: ProjectFile | null,
): ProjectFile[] {
  if (!recovered) return [...diff];
  if (diff.some((f) => f.name === recovered.name)) return [...diff];
  return [...diff, recovered];
}

export async function findSameTurnHtmlWriteForRecoveredArtifact({
  artifactHtml,
  producedFiles,
  readProjectHtml,
}: {
  artifactHtml: string;
  producedFiles: readonly ProjectFile[];
  readProjectHtml: (name: string) => Promise<string | null>;
}): Promise<ProjectFile | null> {
  const recovered = normalizeHtmlForRecoveredArtifactComparison(artifactHtml);
  if (!recovered) return null;
  const candidates = producedFiles.filter(isHtmlProjectFile);
  if (candidates.length === 0) return null;
  const contents = await Promise.all(candidates.map((file) => readProjectHtml(file.name)));
  const normalized = contents.map(normalizeHtmlForRecoveredArtifactComparison);
  // Bind only on an exact normalized-content match. This is inherently
  // agent-agnostic (#4308): whenever a filesystem-backed CLI writes an HTML
  // file and echoes the same document as an artifact, the normalized contents
  // are equal and we suppress the duplicate — no Claude-specific gate needed.
  //
  // We deliberately do NOT bind on a content *mismatch*. A differing same-turn
  // HTML file is a genuinely different document and must persist on its own.
  // A blind single-file bind also mis-fired across queued runs: the pre-turn
  // file snapshot for a queued run can predate the previous run's persist, so
  // computeProducedFiles() reports that earlier artifact as "produced this
  // turn" and we'd bind the echo to the wrong, unrelated file.
  const exact = candidates.find((_file, i) => normalized[i] === recovered);
  return exact ?? null;
}

function isHtmlProjectFile(file: ProjectFile): boolean {
  const name = (file.path || file.name).toLowerCase();
  return file.kind === 'html' || /\.(?:html?|xhtml)$/u.test(name);
}

function normalizeHtmlForRecoveredArtifactComparison(value: string | null | undefined): string {
  return String(value || '')
    .replace(/^\uFEFF/, '')
    .replace(/\r\n?/g, '\n')
    .trim();
}

export function clearStreamingConversationMarker(
  currentConversationId: string | null,
  completedConversationId?: string | null,
): string | null {
  if (
    completedConversationId !== undefined
    && completedConversationId !== null
    && currentConversationId !== completedConversationId
  ) {
    return currentConversationId;
  }
  return null;
}

export function shouldClearActiveRunRefs(
  currentConversationId: string | null,
  completedConversationId: string,
): boolean {
  return currentConversationId === completedConversationId;
}

export function finalizeActiveAssistantMessagesOnStop(
  messages: ChatMessage[],
  stoppedAt: number,
): { messages: ChatMessage[]; finalized: ChatMessage[] } {
  const finalized: ChatMessage[] = [];
  const next = messages.map((message) => {
    if (!isStoppableAssistantMessage(message)) {
      return message;
    }
    const updated = {
      ...message,
      runStatus: 'canceled' as const,
      endedAt: message.endedAt ?? stoppedAt,
    };
    finalized.push(updated);
    return updated;
  });
  return { messages: next, finalized };
}

type BufferedTextUpdates = ReturnType<typeof createBufferedTextUpdates>;

export function createBufferedTextUpdates({
  updateMessage,
  persistSoon,
  flushAndPersistNow,
  onContentDelta,
}: {
  updateMessage: (updater: (prev: ChatMessage) => ChatMessage) => void;
  persistSoon: () => void;
  // Synchronous flush + persist with a transport that survives page
  // unload (PUT with keepalive). Invoked by the pagehide handler so the
  // last buffered chunk isn't lost when the user reloads mid-stream.
  flushAndPersistNow?: () => void;
  onContentDelta?: (delta: string) => void;
}) {
  let pendingContentDelta = '';
  let pendingTextEventDelta = '';
  let flushFrame: number | null = null;
  let flushTimer: ReturnType<typeof setTimeout> | null = null;
  let disposed = false;
  let flushing = false;
  let needsFlush = false;
  const hasDocument = typeof document !== 'undefined';
  const hasWindow = typeof window !== 'undefined';

  const cancelScheduledFlush = () => {
    if (flushFrame !== null) {
      cancelAnimationFrame(flushFrame);
      flushFrame = null;
    }
    if (flushTimer !== null) {
      clearTimeout(flushTimer);
      flushTimer = null;
    }
  };

  const flush = () => {
    if (disposed) return;
    if (flushing) {
      needsFlush = true;
      return;
    }
    cancelScheduledFlush();
    if (!pendingContentDelta && !pendingTextEventDelta && !needsFlush) return;
    flushing = true;
    needsFlush = false;
    const contentDelta = pendingContentDelta;
    const textEventDelta = pendingTextEventDelta;
    pendingContentDelta = '';
    pendingTextEventDelta = '';
    try {
      updateMessage((prev) => ({
        ...prev,
        content: prev.content + contentDelta,
        events: textEventDelta
          ? [...(prev.events ?? []), { kind: 'text', text: textEventDelta }]
          : prev.events,
      }));
      persistSoon();
      if (contentDelta) onContentDelta?.(contentDelta);
    } finally {
      flushing = false;
    }
    if (pendingContentDelta || pendingTextEventDelta || needsFlush) {
      needsFlush = false;
      scheduleFlush();
    }
  };

  const scheduleFlush = () => {
    if (disposed || flushFrame !== null || flushTimer !== null) return;
    flushFrame = requestAnimationFrame(() => {
      flushFrame = null;
      flush();
    });
    flushTimer = setTimeout(() => {
      flushTimer = null;
      flush();
    }, 250);
  };

  const appendContent = (delta: string) => {
    if (disposed) return;
    pendingContentDelta += delta;
    needsFlush = true;
    scheduleFlush();
  };

  const appendTextEvent = (delta: string) => {
    if (disposed) return;
    pendingTextEventDelta += delta;
    needsFlush = true;
    scheduleFlush();
  };

  const appendEvent = (ev: AgentEvent) => {
    if (disposed) return;
    if (ev.kind === 'text') {
      appendTextEvent(ev.text);
      return;
    }
    flush();
    updateMessage((prev) => ({ ...prev, events: [...(prev.events ?? []), ev] }));
    persistSoon();
  };

  const cancel = () => {
    disposed = true;
    cancelScheduledFlush();
    pendingContentDelta = '';
    pendingTextEventDelta = '';
    needsFlush = false;
    if (hasDocument) {
      document.removeEventListener('visibilitychange', onVisibilityChange);
    }
    if (hasWindow) {
      window.removeEventListener('pagehide', onPageHide);
    }
  };

  function onVisibilityChange() {
    if (document.visibilityState === 'hidden') {
      flush();
    }
  }

  function onPageHide() {
    flush();
    // persistSoon's 500ms debounce never fires once the document tears
    // down, so synchronously PUT with keepalive instead.
    flushAndPersistNow?.();
  }

  if (hasDocument) {
    document.addEventListener('visibilitychange', onVisibilityChange);
  }
  if (hasWindow) {
    window.addEventListener('pagehide', onPageHide);
  }

  // True when text has been appended but not yet flushed into a `text` event.
  // Callers that need the soon-to-be-committed event count (e.g. pinning a live
  // tool's stream position) add 1 for this still-buffered preamble.
  const hasPendingText = () => pendingTextEventDelta.length > 0;

  return { appendContent, appendTextEvent, appendEvent, flush, cancel, hasPendingText };
}
