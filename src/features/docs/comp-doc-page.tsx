/**
 * 服务端 Markdown 文档页：读取 `onlyoffice-web-comp/docs/*.md` 并渲染。
 */
import {
  extractDocTitle,
  getMarkdownDocBySlug,
} from "./config/site-map";
import { readCompDoc } from "./lib/server";
import { MarkdownContent } from "./components/markdown-content";

type CompDocPageProps = {
  slug: string;
};

export async function CompDocPage({ slug }: CompDocPageProps) {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) {
    return null;
  }

  const content = await readCompDoc(doc.file);

  return <MarkdownContent content={content} />;
}

export async function getCompDocMetadata(slug: string) {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) return { title: "文档" };

  const content = await readCompDoc(doc.file);
  const title = extractDocTitle(content, doc.label);

  return {
    title: `${title} — OnlyOffice Web Comp`,
  };
}
