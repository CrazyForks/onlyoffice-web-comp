import { DocsShell } from "@/features/docs/components/docs-shell";
import { Suspense } from "react";

export default function DocsLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <Suspense fallback={<div className="px-4 py-8 text-sm text-neutral-500">加载文档…</div>}>
      <DocsShell>{children}</DocsShell>
    </Suspense>
  );
}
