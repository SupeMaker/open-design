// Compact CLI/BYOK execution-mode switch.
//
// Renders a bottom-left toolbar pill ("Local CLI" / "BYOK") that opens a
// dropdown of the two execution modes. Used in the home composer so users can
// choose how the next run executes before pressing Send.

import { useEffect, useRef, useState } from 'react';
import type { ExecMode } from '../types';
import { Icon } from './Icon';
import { useT } from '../i18n';

interface Props {
  mode: ExecMode;
  onChange: (mode: ExecMode) => void;
  onTrackClick?: (modeBefore: ExecMode, modeAfter: ExecMode) => void;
  className?: string;
}

export function ExecutionModePicker({ mode, onChange, onTrackClick, className }: Props) {
  const t = useT();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    function onPointer(event: MouseEvent) {
      if (ref.current?.contains(event.target as Node)) return;
      setOpen(false);
    }
    function onKey(event: KeyboardEvent) {
      if (event.key === 'Escape') setOpen(false);
    }
    document.addEventListener('mousedown', onPointer);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onPointer);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  function select(next: ExecMode) {
    setOpen(false);
    if (next === mode) return;
    onTrackClick?.(mode, next);
    onChange(next);
  }

  return (
    <div
      className={`execution-mode-picker execution-mode-picker--in-input${className ? ` ${className}` : ''}`}
      ref={ref}
    >
      <button
        type="button"
        className="execution-mode-picker__trigger"
        data-testid="execution-mode-trigger"
        aria-expanded={open}
        aria-haspopup="menu"
        title={t('settings.modeAria')}
        onClick={() => setOpen((v) => !v)}
      >
        <Icon name={mode === 'daemon' ? 'terminal' : 'globe'} size={13} />
        <span className="execution-mode-picker__label">
          {mode === 'daemon' ? t('agentPicker.localCli') : t('agentPicker.byok')}
        </span>
        <Icon name="chevron-down" size={11} />
      </button>
      {open ? (
        <div className="execution-mode-picker__menu" role="menu" data-testid="execution-mode-menu">
          <button
            type="button"
            role="menuitem"
            className={`execution-mode-picker__item${mode === 'daemon' ? ' is-active' : ''}`}
            data-testid="execution-mode-daemon"
            onClick={() => select('daemon')}
          >
            <Icon name="terminal" size={13} />
            <span>{t('agentPicker.localCli')}</span>
          </button>
          <button
            type="button"
            role="menuitem"
            className={`execution-mode-picker__item${mode === 'api' ? ' is-active' : ''}`}
            data-testid="execution-mode-api"
            onClick={() => select('api')}
          >
            <Icon name="globe" size={13} />
            <span>{t('agentPicker.byok')}</span>
          </button>
        </div>
      ) : null}
    </div>
  );
}
