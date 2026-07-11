import type { Metadata } from "next";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";
import { OfficePreviewPage } from "@/features/demo/office-preview-page";
import { DocsDemoPage } from "@/features/docs/components/docs-demo-page";
import { readCompDoc } from "@/features/docs/lib/server";

export const metadata: Metadata = {
  title: "单实例示例 — OnlyOffice Web Comp",
};

export default async function DocsDemosSinglePage() {
  const content = await readCompDoc("单实例示例.md");

  return (
    <DocsDemoPage content={content} marker="single">
      <OfficePreviewPage
        embedded
        title="单实例"
        defaultFileName="New_Document.docx"
        fileType={FILE_TYPE.DOCX}
        newButtonLabel="新建文档"
      />
    </DocsDemoPage>
  );
}
