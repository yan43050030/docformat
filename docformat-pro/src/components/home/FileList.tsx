import { X, FileText } from 'lucide-react';

interface FileListProps {
  files: string[];
  onRemove: (index: number) => void;
  className?: string;
}

export function FileList({ files, onRemove, className = '' }: FileListProps) {
  return (
    <div className={`space-y-1 max-h-40 overflow-auto scrollbar-thin ${className}`}>
      {files.map((file, i) => (
        <div
          key={i}
          className="flex items-center justify-between px-3 py-2 rounded-lg bg-[var(--color-paper)] border border-[var(--color-border-light)] text-sm group hover:border-[var(--color-border-medium)] transition-colors"
        >
          <div className="flex items-center gap-2 min-w-0">
            <FileText size={14} className="text-[var(--color-ink-muted)] shrink-0" />
            <span className="truncate text-[var(--color-ink-light)]">{file}</span>
          </div>
          <button
            onClick={() => onRemove(i)}
            className="p-0.5 rounded hover:bg-red-50 text-[var(--color-ink-muted)] hover:text-red-500 cursor-pointer opacity-0 group-hover:opacity-100 transition-all"
          >
            <X size={14} />
          </button>
        </div>
      ))}
    </div>
  );
}
