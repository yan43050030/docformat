interface RadioGroupProps<T extends string> {
  value: T;
  onChange: (value: T) => void;
  options: { label: string; value: T; description?: string }[];
  name: string;
}

export function RadioGroup<T extends string>({ value, onChange, options, name }: RadioGroupProps<T>) {
  return (
    <div className="flex flex-col gap-2">
      {options.map((opt) => (
        <label
          key={opt.value}
          className={`flex items-start gap-3 px-4 py-3 rounded-lg border cursor-pointer transition-all duration-150
            ${value === opt.value
              ? 'border-[var(--color-vermillion)] bg-[var(--color-vermillion)]/5'
              : 'border-[var(--color-border-light)] hover:border-[var(--color-border-medium)] hover:bg-[var(--color-paper-dark)]'
            }`}
        >
          <input
            type="radio"
            name={name}
            value={opt.value}
            checked={value === opt.value}
            onChange={() => onChange(opt.value)}
            className="mt-0.5 accent-[var(--color-vermillion)]"
          />
          <div className="flex flex-col gap-0.5">
            <span className="text-sm font-medium text-[var(--color-ink)]">{opt.label}</span>
            {opt.description && (
              <span className="text-xs text-[var(--color-ink-muted)]">{opt.description}</span>
            )}
          </div>
        </label>
      ))}
    </div>
  );
}
