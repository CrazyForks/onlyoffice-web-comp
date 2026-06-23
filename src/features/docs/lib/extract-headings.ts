import GithubSlugger from "github-slugger";

export type DocHeading = {
  /** 1 = `##`，2 = `###` */
  level: 1 | 2;
  text: string;
  id: string;
};

function stripInlineMarkdown(text: string) {
  return text
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .trim();
}

/** 从 Markdown 提取 `##` / `###`，slug 与 rehype-slug 一致 */
export function extractHeadings(markdown: string): DocHeading[] {
  const slugger = new GithubSlugger();
  const headings: DocHeading[] = [];
  let inCode = false;

  for (const line of markdown.split("\n")) {
    const trimmed = line.trim();
    if (trimmed.startsWith("```")) {
      inCode = !inCode;
      continue;
    }
    if (inCode) continue;

    const h3 = line.match(/^### (?!#)(.+)$/);
    const h2 = line.match(/^## (?!#)(.+)$/);

    if (h3) {
      const text = stripInlineMarkdown(h3[1]);
      headings.push({ level: 2, text, id: slugger.slug(text) });
    } else if (h2) {
      const text = stripInlineMarkdown(h2[1]);
      if (/^(目录|contents?|toc)$/i.test(text)) continue;
      headings.push({ level: 1, text, id: slugger.slug(text) });
    }
  }

  return headings;
}

export function getTocHeadings(markdown: string): DocHeading[] {
  return extractHeadings(markdown);
}
