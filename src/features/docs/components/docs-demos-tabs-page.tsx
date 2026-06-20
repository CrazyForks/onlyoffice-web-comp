"use client";

/**
 * 示例页：按侧栏 `?tab=` 渲染对应 Markdown，并在简介后嵌入 live demo。
 * Markdown 源文件：`onlyoffice-web-comp/docs/08|09-*.md`
 */
import { useSearchParams } from "next/navigation";
import { OfficePreviewPage } from "@/features/demo/office-preview-page";
import { TabsMultiPage } from "@/features/demo/tabs-multi-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";
import type { DemoTabId } from "../config/site-map";
import { MarkdownContent } from "./markdown-content";

type DocsDemosTabsPageProps = {
  contents: Record<DemoTabId, string>;
};

function splitDemoMarkdown(content: string) {
  const [intro, rest = ""] = content.split(/\n## 在线演示\n/);
  const body = rest
    .replace(/<!-- demo:\w+ -->\s*/g, "")
    .replace(/^[^\n]+\n?/, "")
    .trim();

  return { intro: intro.trim(), body };
}

function isDemoTab(value: string | null): DemoTabId {
  return value === "multi" ? "multi" : "single";
}

export function DocsDemosTabsPage({ contents }: DocsDemosTabsPageProps) {
  const searchParams = useSearchParams();
  const activeTab = isDemoTab(searchParams.get("tab"));
  const { intro, body } = splitDemoMarkdown(contents[activeTab]);

  return (
    <div className="w-full max-w-full">
      <MarkdownContent content={intro} />

      <section className="docs-demo-section">
        <h2 className="docs-md-h2">在线演示</h2>
        <div
          className={`docs-demo-shell ${
            activeTab === "single"
              ? "docs-demo-shell--single"
              : "docs-demo-shell--multi"
          }`}
        >
          <div
            className={`absolute inset-0 ${
              activeTab === "single" ? "visible z-10" : "invisible z-0"
            }`}
          >
            <OfficePreviewPage
              embedded
              title="单实例"
              defaultFileName="New_Document.docx"
              fileType={FILE_TYPE.DOCX}
              newButtonLabel="新建文档"
            />
          </div>
          <div
            className={`absolute inset-0 ${
              activeTab === "multi" ? "visible z-10" : "invisible z-0"
            }`}
          >
            <TabsMultiPage embedded />
          </div>
        </div>
      </section>

      {body ? <MarkdownContent content={body} /> : null}
    </div>
  );
}
