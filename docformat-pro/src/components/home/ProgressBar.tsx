interface ProgressBarProps {
  progress: number;
}

export function ProgressBar({ progress }: ProgressBarProps) {
  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between text-sm">
        <span className="font-medium text-[var(--color-ink)]">处理中...</span>
        <span className="text-[var(--color-ink-muted)] tabular-nums">{progress}%</span>
      </div>
      <div className="w-full h-2 bg-[var(--color-paper-dark)] rounded-full overflow-hidden">
        <div
          className="h-full bg-[var(--color-vermillion)] rounded-full transition-all duration-300 ease-out"
          style={{ width: `${progress}%` }}
        />
      </div>
    </div>
  );
}
