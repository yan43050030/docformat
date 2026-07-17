import { useState, useCallback } from 'react';
import { Upload, FileText } from 'lucide-react';

interface FileDropZoneProps {
  selectedFiles: string[];
  onSelectFiles: () => void;
  onRemoveFile: (index: number) => void;
}

export function FileDropZone({ selectedFiles, onSelectFiles }: FileDropZoneProps) {
  const [isDragOver, setIsDragOver] = useState(false);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  if (selectedFiles.length > 0) {
    return (
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <FileText size={18} className="text-[var(--color-vermillion)]" />
          <span className="text-sm font-medium text-[var(--color-ink)]">
            {selectedFiles.length} 个文件已就绪
          </span>
        </div>
        <button
          onClick={onSelectFiles}
          className="text-xs text-[var(--color-teal)] hover:text-[var(--color-teal-light)] font-medium cursor-pointer transition-colors"
        >
          + 添加更多文件
        </button>
      </div>
    );
  }

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      onClick={onSelectFiles}
      className={`relative flex flex-col items-center justify-center py-10 px-4 rounded-xl border-2 border-dashed cursor-pointer transition-all duration-200
        ${isDragOver
          ? 'border-[var(--color-vermillion)] bg-[var(--color-vermillion)]/5'
          : 'border-[var(--color-border-medium)] hover:border-[var(--color-ink-muted)] hover:bg-[var(--color-paper-dark)]/50'
        }`}
    >
      <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 transition-colors
        ${isDragOver ? 'bg-[var(--color-vermillion)]/10' : 'bg-[var(--color-paper-dark)]'}`}
      >
        <Upload size={22} className={isDragOver ? 'text-[var(--color-vermillion)]' : 'text-[var(--color-ink-muted)]'} />
      </div>
      <p className="text-sm font-medium text-[var(--color-ink)]">
        {isDragOver ? '释放文件以添加' : '拖拽文件到此处，或点击选择'}
      </p>
      <p className="text-xs text-[var(--color-ink-muted)] mt-1">
        支持 .docx / .doc / .wps 格式，可多选文件或文件夹
      </p>
    </div>
  );
}
