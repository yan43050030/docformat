import { useAppStore } from '@/lib/store';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Trash2, Info, CheckCircle, AlertTriangle, AlertCircle } from 'lucide-react';

const levelIcons = {
  info: Info,
  success: CheckCircle,
  warning: AlertTriangle,
  error: AlertCircle,
};

const levelBadgeVariant: Record<string, 'default' | 'success' | 'warning' | 'error'> = {
  info: 'default',
  success: 'success',
  warning: 'warning',
  error: 'error',
};

export function LogPage() {
  const { logs, clearLogs } = useAppStore();

  return (
    <div className="max-w-3xl mx-auto px-6 py-8 space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-xl font-bold text-[var(--color-ink)]">处理日志</h1>
          <p className="text-sm text-[var(--color-ink-muted)] mt-1">记录所有文件处理操作的详细日志</p>
        </div>
        {logs.length > 0 && (
          <Button variant="ghost" size="sm" onClick={clearLogs}>
            <Trash2 size={14} /> 清空日志
          </Button>
        )}
      </div>

      {logs.length === 0 ? (
        <Card className="text-center py-16">
          <Info size={36} className="mx-auto text-[var(--color-border-medium)] mb-3" />
          <p className="text-sm text-[var(--color-ink-muted)]">暂无处理日志</p>
          <p className="text-xs text-[var(--color-ink-muted)] mt-1">处理文件后，日志将显示在这里</p>
        </Card>
      ) : (
        <Card padding="none">
          <div className="divide-y divide-[var(--color-border-light)]">
            {logs.map((log) => {
              const Icon = levelIcons[log.level];
              return (
                <div key={log.id} className="flex items-start gap-3 px-4 py-3 hover:bg-[var(--color-paper)] transition-colors">
                  <Icon
                    size={16}
                    className={
                      log.level === 'success' ? 'text-[var(--color-success)] mt-0.5' :
                      log.level === 'error' ? 'text-[var(--color-error)] mt-0.5' :
                      log.level === 'warning' ? 'text-[var(--color-warning)] mt-0.5' :
                      'text-[var(--color-ink-muted)] mt-0.5'
                    }
                  />
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="text-sm text-[var(--color-ink)]">{log.message}</span>
                      <Badge variant={levelBadgeVariant[log.level]}>
                        {log.level === 'info' ? '信息' :
                         log.level === 'success' ? '成功' :
                         log.level === 'warning' ? '警告' : '错误'}
                      </Badge>
                    </div>
                    <p className="text-[10px] text-[var(--color-ink-muted)] mt-0.5">
                      {new Date(log.timestamp).toLocaleString('zh-CN')}
                    </p>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}
    </div>
  );
}
