"use client";

/**
 * 主页 Hero 区 Office 文档窗口动效（DOCX / XLSX / PPTX 浮动卡片）。
 */
import { BrandMark } from "@/features/shell";

const FLOATERS = [
  { label: "DOCX", x: "8%", y: "14%", delay: "0s", w: "w-[42%]" },
  { label: "XLSX", x: "52%", y: "8%", delay: "0.6s", w: "w-[38%]" },
  { label: "PPTX", x: "28%", y: "48%", delay: "1.1s", w: "w-[44%]" },
] as const;

function DocWindow({
  label,
  className = "",
}: {
  label: string;
  className?: string;
}) {
  return (
    <div
      className={`border border-neutral-900 bg-white shadow-[6px_6px_0_0_#171717] ${className}`}
    >
      <div className="flex items-center gap-1.5 border-b border-neutral-200 px-2 py-1.5">
        <span className="h-1.5 w-1.5 rounded-full bg-neutral-400" />
        <span className="h-1.5 w-1.5 rounded-full bg-neutral-300" />
        <span className="h-1.5 w-1.5 rounded-full bg-neutral-300" />
        <span className="ml-auto text-[9px] font-medium tracking-wide text-neutral-500">
          {label}
        </span>
      </div>
      <div className="space-y-1.5 p-2.5">
        <div className="h-1 w-[88%] bg-neutral-900" />
        <div className="h-1 w-[72%] bg-neutral-300" />
        <div className="h-1 w-[80%] bg-neutral-200" />
        <div className="mt-2 grid grid-cols-3 gap-1">
          <div className="h-6 border border-neutral-200 bg-neutral-50" />
          <div className="h-6 border border-neutral-200 bg-neutral-50" />
          <div className="h-6 border border-neutral-200 bg-neutral-50" />
        </div>
      </div>
    </div>
  );
}

export function OfficeHeroVisual() {
  return (
    <div className="office-hero-stage relative aspect-[5/4] w-full overflow-hidden border border-neutral-900 bg-[#f7f7f5]">
      <div
        className="absolute inset-0 opacity-[0.35]"
        style={{
          backgroundImage:
            "linear-gradient(#d4d4d4 1px, transparent 1px), linear-gradient(90deg, #d4d4d4 1px, transparent 1px)",
          backgroundSize: "24px 24px",
        }}
      />

      <div className="absolute left-4 top-4 flex items-center gap-2 border border-neutral-300 bg-white/90 px-2 py-1 backdrop-blur-sm">
        <BrandMark className="h-5 w-5" />
        <span className="text-[10px] font-medium tracking-[0.08em] text-neutral-600">
          OFFICE SDK
        </span>
      </div>

      {FLOATERS.map((item) => (
        <div
          key={item.label}
          className={`office-hero-float absolute ${item.w}`}
          style={{
            left: item.x,
            top: item.y,
            animationDelay: item.delay,
          }}
        >
          <DocWindow label={item.label} />
        </div>
      ))}

      <div className="office-hero-scan pointer-events-none absolute inset-x-0 top-0 h-px bg-neutral-900/20" />
    </div>
  );
}
