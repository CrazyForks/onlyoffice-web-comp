/**
 * 文档站点唯一配置：Markdown 条目、示例 Tab、侧栏分组与 slug 解析。
 * 新增文档页只需在此追加 `DOC_ENTRIES` 或 `DEMO_TABS` 条目。
 */
export type DocsSection = "intro" | "demo" | "guide" | "reference";

export type DemoTabId = "single" | "multi";

export type MarkdownDoc = {
  kind: "markdown";
  slug: string;
  file: string;
  label: string;
  href: string;
  section: Exclude<DocsSection, "demo">;
};

export type DocsNavItem = {
  href: string;
  label: string;
};

export type DocsNavGroup = {
  title: string;
  items: DocsNavItem[];
};

/** 示例 Tab：Markdown 源文件 + 站点 query 路由 */
export const DEMO_TABS: {
  id: DemoTabId;
  file: string;
  label: string;
  href: string;
}[] = [
  {
    id: "single",
    file: "08-单实例示例.md",
    label: "单实例",
    href: "/docs/demos?tab=single",
  },
  {
    id: "multi",
    file: "09-多实例示例.md",
    label: "多实例",
    href: "/docs/demos?tab=multi",
  },
];

export const DOC_SECTIONS: { id: DocsSection; title: string }[] = [
  { id: "intro", title: "入门" },
  { id: "demo", title: "示例" },
  { id: "guide", title: "集成指南" },
  { id: "reference", title: "参考" },
];

export const DOC_ENTRIES: MarkdownDoc[] = [
  {
    kind: "markdown",
    slug: "",
    file: "00-概述.md",
    label: "概述",
    href: "/docs",
    section: "intro",
  },
  {
    kind: "markdown",
    slug: "getting-started",
    file: "01-快速开始.md",
    label: "快速开始",
    href: "/docs/getting-started",
    section: "intro",
  },
  {
    kind: "markdown",
    slug: "api",
    file: "02-核心API.md",
    label: "核心 API",
    href: "/docs/api",
    section: "guide",
  },
  {
    kind: "markdown",
    slug: "events",
    file: "03-事件系统.md",
    label: "事件系统",
    href: "/docs/events",
    section: "guide",
  },
  {
    kind: "markdown",
    slug: "example",
    file: "04-完整示例.md",
    label: "完整示例",
    href: "/docs/example",
    section: "guide",
  },
  {
    kind: "markdown",
    slug: "reference",
    file: "05-API参考.md",
    label: "API 参考",
    href: "/docs/reference",
    section: "reference",
  },
  {
    kind: "markdown",
    slug: "formats",
    file: "06-注意事项与支持格式.md",
    label: "注意事项与格式",
    href: "/docs/formats",
    section: "reference",
  },
  {
    kind: "markdown",
    slug: "fonts",
    file: "10-字体配置.md",
    label: "字体配置",
    href: "/docs/fonts",
    section: "reference",
  },
  {
    kind: "markdown",
    slug: "word-api",
    file: "07-批注修订与-Word-API.md",
    label: "批注与 Word API",
    href: "/docs/word-api",
    section: "reference",
  },
];

export const MARKDOWN_DOCS = DOC_ENTRIES;

const FILE_TO_HREF = Object.fromEntries([
  ...MARKDOWN_DOCS.map((doc) => [doc.file, doc.href]),
  ...DEMO_TABS.map((tab) => [tab.file, tab.href]),
]) as Record<string, string>;

export function getDocsNav(): DocsNavGroup[] {
  return DOC_SECTIONS.map((section) => {
    if (section.id === "demo") {
      return {
        title: section.title,
        items: DEMO_TABS.map((tab) => ({
          href: tab.href,
          label: tab.label,
        })),
      };
    }

    return {
      title: section.title,
      items: MARKDOWN_DOCS.filter((entry) => entry.section === section.id).map(
        (entry) => ({
          href: entry.href,
          label: entry.label,
        }),
      ),
    };
  }).filter((group) => group.items.length > 0);
}

export function getMarkdownDocBySlug(slug: string): MarkdownDoc | undefined {
  return MARKDOWN_DOCS.find((doc) => doc.slug === slug);
}

export function getMarkdownSlugs(): string[] {
  return MARKDOWN_DOCS.filter((doc) => doc.slug).map((doc) => doc.slug);
}

export function resolveMarkdownHref(
  href: string | undefined,
): string | undefined {
  if (!href) return href;

  const mdMatch = href.match(/(?:\.\/)?(\d{2}-[^/]+\.md)$/);
  if (mdMatch) {
    const mapped = FILE_TO_HREF[mdMatch[1]];
    if (mapped) return mapped;
  }

  if (href.startsWith("../readme")) {
    return "https://github.com";
  }

  return href;
}

export function isDocsNavActive(
  pathname: string,
  href: string,
  searchParams?: URLSearchParams,
): boolean {
  if (href.startsWith("/docs/demos")) {
    if (pathname !== "/docs/demos") return false;
    const tab = searchParams?.get("tab") ?? "single";
    const hrefTab = new URL(href, "http://localhost").searchParams.get("tab");
    return tab === (hrefTab ?? "single");
  }

  if (href === "/docs") return pathname === "/docs";
  return pathname === href;
}

export function isDocsPath(pathname: string) {
  return pathname === "/docs" || pathname.startsWith("/docs/");
}

export function extractDocTitle(content: string, fallback: string): string {
  const match = content.match(/^#\s+(.+)$/m);
  if (!match) return fallback;
  return match[1].replace(/^\d{2}\s*-\s*/, "").trim();
}
