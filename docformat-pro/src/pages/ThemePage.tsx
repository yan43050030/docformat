import { useAppStore, type ThemeId } from '@/lib/store';
import { Card } from '@/components/ui/card';
import { Palette, Check } from 'lucide-react';

const themes: { id: ThemeId; name: string; description: string; colors: string[]; preview: string }[] = [
  {
    id: 'paper',
    name: '米白纸张',
    description: '温暖舒适的纸张质感，适合长时间办公使用',
    colors: ['#FBF9F6', '#BC4B26', '#2F6F73', '#2E2E2E'],
    preview: 'bg-[#FBF9F6]',
  },
  {
    id: 'dark',
    name: '暗夜模式',
    description: '深色界面，减少眼部疲劳，适合夜间工作',
    colors: ['#1A1A1A', '#D4663F', '#3A8A8F', '#E8E4DF'],
    preview: 'bg-[#1A1A1A]',
  },
  {
    id: 'ink',
    name: '墨韵',
    description: '传统水墨风格，典雅沉稳的东方美学',
    colors: ['#F5F0EB', '#2E2E2E', '#5B7B6C', '#1A1A1A'],
    preview: 'bg-[#F5F0EB]',
  },
  {
    id: 'teal',
    name: '青瓷',
    description: '清新淡雅的青绿色调，现代简约风格',
    colors: ['#F0F7F6', '#2F6F73', '#BC4B26', '#2E2E2E'],
    preview: 'bg-[#F0F7F6]',
  },
];

export function ThemePage() {
  const { theme, setTheme } = useAppStore();

  return (
    <div className="max-w-3xl mx-auto px-6 py-8 space-y-6">
      <div>
        <h1 className="text-xl font-bold text-[var(--color-ink)]">主题设置</h1>
        <p className="text-sm text-[var(--color-ink-muted)] mt-1">选择适合您的界面配色方案</p>
      </div>

      <div className="grid grid-cols-2 gap-4">
        {themes.map((t) => (
          <Card
            key={t.id}
            className={`cursor-pointer transition-all duration-200 hover:shadow-md ${
              theme === t.id ? 'ring-2 ring-[var(--color-vermillion)]' : ''
            }`}
            onClick={() => setTheme(t.id)}
          >
            {/* Color Preview Bar */}
            <div className="flex gap-1 mb-3">
              {t.colors.map((c, i) => (
                <div
                  key={i}
                  className="flex-1 h-8 rounded-md"
                  style={{ backgroundColor: c }}
                />
              ))}
            </div>

            <div className="flex items-start justify-between">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-semibold text-[var(--color-ink)]">{t.name}</span>
                  {theme === t.id && (
                    <span className="w-4 h-4 rounded-full bg-[var(--color-teal)] flex items-center justify-center">
                      <Check size={10} className="text-white" />
                    </span>
                  )}
                </div>
                <p className="text-xs text-[var(--color-ink-muted)] mt-1">{t.description}</p>
              </div>
              <Palette size={16} className="text-[var(--color-ink-muted)]" />
            </div>
          </Card>
        ))}
      </div>

      <Card>
        <h3 className="text-sm font-semibold text-[var(--color-ink)] mb-2">关于主题</h3>
        <p className="text-xs text-[var(--color-ink-muted)] leading-relaxed">
          主题仅影响软件界面的显示效果，不会改变文档排版的实际输出。所有排版输出均严格按照您选择的预设方案执行。
          主题切换会立即生效，无需重新启动应用。
        </p>
      </Card>
    </div>
  );
}
