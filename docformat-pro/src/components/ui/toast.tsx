import { useState, createContext, useContext, useCallback, type ReactNode } from 'react';
import { X, CheckCircle, AlertCircle, AlertTriangle, Info } from 'lucide-react';

type ToastLevel = 'info' | 'success' | 'warning' | 'error';

interface Toast {
  id: string;
  level: ToastLevel;
  message: string;
}

interface ToastContextType {
  toast: (level: ToastLevel, message: string) => void;
}

const ToastContext = createContext<ToastContextType>({ toast: () => {} });

export function useToast() {
  return useContext(ToastContext);
}

const icons: Record<ToastLevel, typeof CheckCircle> = {
  info: Info,
  success: CheckCircle,
  warning: AlertTriangle,
  error: AlertCircle,
};

const colors: Record<ToastLevel, string> = {
  info: 'border-[var(--color-teal)] bg-[var(--color-teal-dim)]',
  success: 'border-green-400 bg-green-50',
  warning: 'border-yellow-400 bg-yellow-50',
  error: 'border-[var(--color-vermillion)] bg-[var(--color-vermillion-dim)]',
};

const iconColors: Record<ToastLevel, string> = {
  info: 'text-[var(--color-teal)]',
  success: 'text-green-600',
  warning: 'text-yellow-600',
  error: 'text-[var(--color-vermillion)]',
};

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const addToast = useCallback((level: ToastLevel, message: string) => {
    const id = crypto.randomUUID();
    setToasts((prev) => [...prev, { id, level, message }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 3000);
  }, []);

  return (
    <ToastContext.Provider value={{ toast: addToast }}>
      {children}
      <div className="fixed bottom-4 right-4 z-50 flex flex-col gap-2 pointer-events-none">
        {toasts.map((t) => {
          const Icon = icons[t.level];
          return (
            <div
              key={t.id}
              className={`flex items-center gap-2 px-4 py-3 rounded-lg border shadow-lg pointer-events-auto animate-slide-up
                ${colors[t.level]} min-w-72 max-w-96`}
            >
              <Icon size={16} className={iconColors[t.level]} />
              <span className="text-sm flex-1 text-[var(--color-ink)]">{t.message}</span>
              <button
                onClick={() => setToasts((prev) => prev.filter((x) => x.id !== t.id))}
                className="p-0.5 rounded hover:bg-black/5 cursor-pointer"
              >
                <X size={14} className="text-[var(--color-ink-muted)]" />
              </button>
            </div>
          );
        })}
      </div>
    </ToastContext.Provider>
  );
}
