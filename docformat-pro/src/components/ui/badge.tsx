interface BadgeProps {
  children: React.ReactNode;
  variant?: 'default' | 'success' | 'warning' | 'error';
  className?: string;
}

const variantStyles: Record<string, string> = {
  default: 'bg-[var(--color-paper-dark)] text-[var(--color-ink-light)]',
  success: 'bg-[var(--color-teal-dim)] text-[var(--color-teal)]',
  warning: 'bg-yellow-50 text-yellow-700',
  error: 'bg-[var(--color-vermillion-dim)] text-[var(--color-vermillion)]',
};

export function Badge({ children, variant = 'default', className = '' }: BadgeProps) {
  return (
    <span className={`inline-flex items-center px-2 py-0.5 text-xs font-medium rounded-full ${variantStyles[variant]} ${className}`}>
      {children}
    </span>
  );
}
