import type { ButtonHTMLAttributes, ReactNode, SelectHTMLAttributes } from "react";

export const demoHeaderClass =
  "border-b border-neutral-200/80 bg-[#fbfbfa]";

export const demoHeaderInnerClass =
  "flex flex-wrap items-center gap-x-4 gap-y-2 px-4 py-2.5";

export const demoToolbarClass = "flex flex-wrap items-center gap-1.5";

export const demoTitleClass = "text-[15px] font-medium tracking-[-0.01em] text-neutral-900";

export const demoSubtitleClass = "truncate text-[12px] text-neutral-500";

const buttonBase =
  "inline-flex h-8 shrink-0 items-center justify-center border px-2.5 text-[13px] leading-none text-neutral-700 disabled:cursor-not-allowed disabled:opacity-40";

export const demoButtonClass = `${buttonBase} border-neutral-300 bg-white hover:border-neutral-400 hover:bg-neutral-50 active:bg-neutral-100`;

export const demoButtonActiveClass = `${buttonBase} border-neutral-400 bg-neutral-100 text-neutral-900`;

type DemoButtonProps = ButtonHTMLAttributes<HTMLButtonElement> & {
  active?: boolean;
};

export function DemoButton({ active, className = "", ...props }: DemoButtonProps) {
  return (
    <button
      type="button"
      className={`${active ? demoButtonActiveClass : demoButtonClass} ${className}`}
      {...props}
    />
  );
}

type DemoFieldProps = {
  label: string;
  children: ReactNode;
  className?: string;
};

export function DemoField({ label, children, className = "" }: DemoFieldProps) {
  return (
    <label
      className={`inline-flex h-8 items-center gap-1.5 border border-neutral-300 bg-white px-2 text-[13px] text-neutral-600 ${className}`}
    >
      <span className="shrink-0 select-none whitespace-nowrap">{label}</span>
      {children}
    </label>
  );
}

type DemoSelectProps = SelectHTMLAttributes<HTMLSelectElement>;

export function DemoSelect({ className = "", ...props }: DemoSelectProps) {
  return (
    <select
      className={`h-6 min-w-0 border-0 bg-transparent py-0 pr-5 pl-0.5 text-[13px] text-neutral-800 outline-none ${className}`}
      {...props}
    />
  );
}

type DemoMenuProps = {
  label: string;
  children: ReactNode;
  disabled?: boolean;
};

export function DemoMenu({ label, children, disabled = false }: DemoMenuProps) {
  return (
    <details className="group relative shrink-0">
      <summary
        aria-disabled={disabled}
        className={`${demoButtonClass} list-none cursor-pointer select-none marker:hidden group-open:border-neutral-400 group-open:bg-neutral-100 ${
          disabled ? "pointer-events-none opacity-40" : ""
        }`}
      >
        {label}
      </summary>
      <div className="absolute right-0 top-[calc(100%+6px)] z-50 w-max max-w-[calc(100vw-24px)] border border-neutral-200 bg-white p-2 shadow-lg">
        <div className="flex flex-col gap-2">{children}</div>
      </div>
    </details>
  );
}

export function DemoMenuRow({ children }: { children: ReactNode }) {
  return <div className="flex flex-wrap items-center gap-1.5">{children}</div>;
}
