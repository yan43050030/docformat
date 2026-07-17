import type { ButtonHTMLAttributes, ReactNode } from 'react';

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger';
  size?: 'sm' | 'md' | 'lg';
  children: ReactNode;
}

const variants: Record<string, string> = {
  primary: 'bg-[var(--color-vermillion)] text-white hover:bg-[var(--color-vermillion-light)] shadow-sm',
  secondary: 'bg-white border border-[var(--color-border-medium)] text-[var(--color-ink)] hover:bg-[var(--color-paper-dark)]',
  ghost: 'text-[var(--color-ink-light)] hover:bg-[var(--color-paper-dark)] hover:text-[var(--color-ink)]',
  danger: 'bg-red-50 text-red-600 hover:bg-red-100 border border-red-200',
};

const sizes: Record<string, string> = {
  sm: 'px-3 py-1.5 text-xs rounded-md',
  md: 'px-4 py-2 text-sm rounded-lg',
  lg: 'px-6 py-3 text-base rounded-lg',
};

export function Button({ variant = 'primary', size = 'md', className = '', children, disabled, ...props }: ButtonProps) {
  return (
    <button
      className={`inline-flex items-center justify-center gap-2 font-medium transition-all duration-150 cursor-pointer
        ${variants[variant]} ${sizes[size]}
        ${disabled ? 'opacity-50 cursor-not-allowed' : ''}
        ${className}`}
      disabled={disabled}
      {...props}
    >
      {children}
    </button>
  );
}
