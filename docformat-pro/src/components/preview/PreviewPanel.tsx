import { useAppStore } from '@/lib/store';
import { Card } from '@/components/ui/card';
import { FileText, EyeOff, Eye } from 'lucide-react';

export function PreviewPanel() {
  const { showPreview, setShowPreview, selectedFiles } = useAppStore();

  if (!showPreview) {
    return (
      <button
        onClick={() => setShowPreview(true)}
        className="fixed right-4 top-4 p-2 rounded-lg bg-white border border-[var(--color-border-light)] shadow-sm hover:bg-[var(--color-paper)] cursor-pointer transition-all z-10"
        title="显示预览"
      >
        <Eye size={16} className="text-[var(--color-ink-muted)]" />
      </button>
    );
  }

  return (
    <aside className="w-80 shrink-0 bg-white border-l border-[var(--color-border-light)] flex flex-col">
      <div className="h-10 flex items-center justify-between px-3 border-b border-[var(--color-border-light)]">
        <span className="text-xs font-medium text-[var(--color-ink-light)]">文档预览</span>
        <button
          onClick={() => setShowPreview(false)}
          className="p-1 rounded hover:bg-[var(--color-paper)] cursor-pointer transition-colors"
          title="隐藏预览"
        >
          <EyeOff size={14} className="text-[var(--color-ink-muted)]" />
        </button>
      </div>

      <div className="flex-1 overflow-auto scrollbar-thin p-4">
        {selectedFiles.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-full text-center">
            <FileText size={40} className="text-[var(--color-border-medium)] mb-3" />
            <p className="text-xs text-[var(--color-ink-muted)] leading-relaxed">
              选择文件后，此处将显示文档的实时预览效果
            </p>
            <p className="text-[10px] text-[var(--color-ink-muted)] mt-1">
              预览模拟当前预设方案的排版效果
            </p>
          </div>
        ) : (
          <div className="space-y-3">
            <Card padding="sm" className="bg-[var(--color-paper)] border-dashed">
              <p className="text-[10px] text-[var(--color-ink-muted)] text-center">
                预览功能将在完整版中启用
              </p>
              <p className="text-xs text-[var(--color-ink)] text-center mt-1 font-medium">
                {selectedFiles[0]}
              </p>
            </Card>
            <div className="space-y-2">
              <PreviewSkeleton />
            </div>
          </div>
        )}
      </div>
    </aside>
  );
}

function PreviewSkeleton() {
  return (
    <div className="space-y-2 animate-pulse">
      <div className="h-4 bg-[var(--color-paper-dark)] rounded w-3/4 mx-auto" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-full" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-full" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-11/12" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-full" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-4/5" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-full" />
      <div className="h-2 bg-[var(--color-paper-dark)] rounded w-3/4" />
    </div>
  );
}
