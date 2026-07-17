import { type InputHTMLAttributes } from 'react';

interface InputProps extends InputHTMLAttributes<HTMLInputElement> {
  label?: string;
}

export function Input({ label, className = '', id, ...props }: InputProps) {
  return (
    <div className="flex flex-col gap-1.5">
      {label && (
        <label htmlFor={id} className="text-xs font-medium text-[var(--color-ink-light)]">
          {label}
        </label>
      )}
      <input
        id={id}
        className={`w-full px-3 py-2 text-sm bg-white border border-[var(--color-border-medium)] rounded-lg
          placeholder:text-[var(--color-ink-muted)]
          focus:outline-none focus:ring-2 focus:ring-[var(--color-vermillion)]/20 focus:border-[var(--color-vermillion)]
          transition-all duration-150 ${className}`}
        {...props}
      />
    </div>
  );
}
