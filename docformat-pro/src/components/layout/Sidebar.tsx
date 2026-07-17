import { useAppStore } from '@/lib/store';
import { FileText, Settings, Palette, ScrollText } from 'lucide-react';

const navItems = [
  { id: 'home' as const, label: '处理', icon: FileText },
  { id: 'presets' as const, label: '预设方案', icon: Settings },
  { id: 'theme' as const, label: '主题', icon: Palette },
  { id: 'log' as const, label: '日志', icon: ScrollText },
];

export function Sidebar() {
  const { activePage, setActivePage } = useAppStore();

  return (
    <aside className="w-56 shrink-0 bg-white border-r border-[var(--color-border-light)] flex flex-col">
      <div className="h-14 flex items-center px-5 border-b border-[var(--color-border-light)]">
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded-lg bg-[var(--color-vermillion)] flex items-center justify-center">
            <FileText size={14} className="text-white" />
          </div>
          <span className="text-sm font-bold text-[var(--color-ink)] tracking-tight">DocFormat</span>
          <span className="text-[10px] font-medium text-[var(--color-vermillion)] bg-[var(--color-vermillion-dim)] px-1.5 py-0.5 rounded">Pro</span>
        </div>
      </div>

      <nav className="flex-1 px-3 py-4 space-y-1">
        {navItems.map(({ id, label, icon: Icon }) => {
          const isActive = activePage === id;
          return (
            <button
              key={id}
              onClick={() => setActivePage(id)}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all duration-150 cursor-pointer
                ${isActive
                  ? 'bg-[var(--color-vermillion)]/8 text-[var(--color-vermillion)]'
                  : 'text-[var(--color-ink-light)] hover:bg-[var(--color-paper-dark)] hover:text-[var(--color-ink)]'
                }`}
            >
              <Icon size={18} strokeWidth={1.8} />
              {label}
            </button>
          );
        })}
      </nav>

      <div className="px-4 py-3 border-t border-[var(--color-border-light)]">
        <div className="text-[10px] text-[var(--color-ink-muted)] uppercase tracking-wider">版本 1.0.0</div>
      </div>
    </aside>
  );
}
