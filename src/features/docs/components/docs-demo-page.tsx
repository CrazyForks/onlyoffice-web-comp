"use client";

import type { ReactNode } from "react";
import { useMemo } from "react";
import { getTocHeadings } from "../lib/extract-headings";
import { DocToc } from "./doc-toc";
import { MarkdownContent } from "./markdown-content";

type DocsDemoPageProps = {
  content: string;
  marker: "single" | "multi";
  children: ReactNode;
};

function splitDemoMarkdown(content: string, marker: string) {
  const slot = `<!-- demo:${marker} -->`;
  const index = content.indexOf(slot);
  if (index < 0) {
    return { before: content.trim(), after: "" };
  }

  return {
    before: content.slice(0, index).trim(),
    after: content.slice(index + slot.length).trim(),
  };
}

export function DocsDemoPage({ content, marker, children }: DocsDemoPageProps) {
  const { before, after } = useMemo(
    () => splitDemoMarkdown(content, marker),
    [content, marker],
  );
  const tocHeadings = useMemo(() => getTocHeadings(content), [content]);

  return (
    <div className="docs-article-layout docs-demo-article-layout w-full max-w-full">
      <div className="min-w-0">
        <MarkdownContent content={before} showToc={false} />
        <section
          className={`docs-demo-shell docs-demo-shell--${marker}`}
          aria-label={marker === "single" ? "单实例在线演示" : "多实例在线演示"}
        >
          {children}
        </section>
        {after ? <MarkdownContent content={after} showToc={false} /> : null}
      </div>
      <DocToc headings={tocHeadings} />
    </div>
  );
}
