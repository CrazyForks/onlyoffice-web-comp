import type { Metadata } from "next";
import { TabsMultiPage } from "@/features/demo/tabs-multi-page";
import { DocsDemoPage } from "@/features/docs/components/docs-demo-page";
import { readCompDoc } from "@/features/docs/lib/server";

export const metadata: Metadata = {
  title: "多实例示例 — OnlyOffice Web Comp",
};

export default async function DocsDemosMultiPage() {
  const content = await readCompDoc("多实例示例.md");

  return (
    <DocsDemoPage content={content} marker="multi">
      <TabsMultiPage embedded />
    </DocsDemoPage>
  );
}
