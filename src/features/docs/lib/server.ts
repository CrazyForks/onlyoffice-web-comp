import "server-only";

/**
 * 服务端读取 `onlyoffice-web-comp/docs/` 下的 Markdown 源文件。
 */
import fs from "fs/promises";
import path from "path";
import { getMarkdownDocBySlug } from "../config/site-map";

export const COMP_DOCS_DIR = path.join(
  process.cwd(),
  "src/components/onlyoffice-web-comp/docs",
);

export async function readCompDoc(file: string): Promise<string> {
  const filePath = path.join(COMP_DOCS_DIR, file);
  return fs.readFile(filePath, "utf-8");
}

export async function readMarkdownDocBySlug(slug: string): Promise<string | null> {
  const doc = getMarkdownDocBySlug(slug);
  if (!doc) return null;
  return readCompDoc(doc.file);
}
