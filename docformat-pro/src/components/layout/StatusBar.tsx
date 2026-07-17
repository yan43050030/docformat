import { useAppStore } from '@/lib/store';
import { CheckCircle, AlertCircle, Clock } from 'lucide-react';

export function StatusBar() {
  const { presets, activePresetId, logs } = useAppStore();
  const activePreset = presets.find((p) => p.id === activePresetId);
  const recentLogs = logs.slice(0, 3);

  return (
    <footer className="h-8 shrink-0 bg-white border-t border-[var(--color-border-light)] flex items-center justify-between px-4">
      <div className="flex items-center gap-3 text-xs text-[var(--color-ink-muted)]">
        <span className="flex items-center gap-1">
          <Clock size={12} />
          当前预设: <span className="font-medium text-[var(--color-ink-light)]">{activePreset?.name}</span>
        </span>
        {recentLogs.length > 0 && (
          <span className="flex items-center gap-1">
            {recentLogs[0].level === 'success' ? <CheckCircle size={12} className="text-[var(--color-success)]" />
              : recentLogs[0].level === 'error' ? <AlertCircle size={12} className="text-[var(--color-error)]" />
              : <Clock size={12} />}
            {recentLogs[0].message}
          </span>
        )}
      </div>
      <div className="text-[10px] text-[var(--color-ink-muted)]">
        v1.0.0 · GB/T 9704-2012
      </div>
    </footer>
  );
}
