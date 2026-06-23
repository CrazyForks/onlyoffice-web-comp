import { Suspense } from "react";
import type { Metadata } from "next";
import { DocsDemosTabsPage } from "@/features/docs/components/docs-demos-tabs-page";
import { DEMO_TABS, type DemoTabId } from "@/features/docs/config/site-map";
import { readCompDoc } from "@/features/docs/lib/server";

export const metadata: Metadata = {
  title: "在线示例 — OnlyOffice Web Comp",
};

async function loadDemoContents(): Promise<Record<DemoTabId, string>> {
  const entries = await Promise.all(
    DEMO_TABS.map(async (tab) => [tab.id, await readCompDoc(tab.file)] as const),
  );
  return Object.fromEntries(entries) as Record<DemoTabId, string>;
}

export default async function DocsDemosRoutePage() {
  const contents = await loadDemoContents();

  return (
    <Suspense fallback={<div className="docs-md-p">加载示例…</div>}>
      <DocsDemosTabsPage contents={contents} />
    </Suspense>
  );
}
