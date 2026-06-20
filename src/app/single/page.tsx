import type { Metadata } from "next";
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";

export const metadata: Metadata = {
  title: "单实例 — OnlyOffice MVP",
};

export default function SingleInstancePage() {
  return (
    <OfficePreviewPage
      title="单实例演示"
      defaultFileName="New_Document.docx"
      fileType={FILE_TYPE.DOCX}
      accept=".docx,.doc,.odt,.rtf,.txt,.docm"
      newButtonLabel="新建文档"
    />
  );
}
