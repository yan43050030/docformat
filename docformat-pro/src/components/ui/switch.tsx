interface SwitchProps {
  checked: boolean;
  onChange: (checked: boolean) => void;
  label?: string;
  id?: string;
}

export function Switch({ checked, onChange, label, id }: SwitchProps) {
  return (
    <label htmlFor={id} className="flex items-center gap-2 cursor-pointer select-none">
      <button
        id={id}
        type="button"
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative inline-flex h-5 w-9 shrink-0 items-center rounded-full transition-all duration-200 cursor-pointer
          ${checked ? 'bg-[var(--color-teal)]' : 'bg-[var(--color-border-medium)]'}`}
      >
        <span
          className={`inline-block h-4 w-4 rounded-full bg-white shadow-sm transition-transform duration-200
            ${checked ? 'translate-x-[18px]' : 'translate-x-[2px]'}`}
        />
      </button>
      {label && <span className="text-sm text-[var(--color-ink)]">{label}</span>}
    </label>
  );
}
