import type { Metadata } from "next";
import Link from "next/link";
import { DEMO_TABS } from "@/features/docs/config/site-map";

export const metadata: Metadata = {
  title: "在线示例 — OnlyOffice Web Comp",
};

export default function DocsDemosRoutePage() {
  return (
    <div className="max-w-3xl">
      <h1 className="docs-md-h1">在线示例</h1>
      <p className="docs-md-p">选择一个示例路由打开对应的编辑器实例。</p>
      <div className="mt-6 grid gap-3 sm:grid-cols-2">
        {DEMO_TABS.map((tab) => (
          <Link
            key={tab.id}
            href={tab.href}
            className="block border border-neutral-200 bg-white p-4 text-neutral-900 transition-colors hover:border-neutral-400 hover:bg-neutral-50"
          >
            <span className="text-[15px] font-medium">{tab.label}</span>
            <span className="mt-1 block text-[13px] text-neutral-500">
              {tab.id === "single"
                ? "一个页面只挂载一个编辑器实例"
                : "多个 Tab 按 containerId 隔离实例"}
            </span>
          </Link>
        ))}
      </div>
    </div>
  );
}
