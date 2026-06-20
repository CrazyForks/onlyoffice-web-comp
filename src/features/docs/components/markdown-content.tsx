"use client";

/**
 * 将组件库 Markdown 渲染为文档站 HTML，并改写 `./xx.md` 为站内路由。
 */
import Link from "next/link";
import type { Components } from "react-markdown";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { resolveMarkdownHref } from "../config/site-map";

const components: Components = {
  h1: ({ children }) => <h1 className="docs-md-h1">{children}</h1>,
  h2: ({ children }) => <h2 className="docs-md-h2">{children}</h2>,
  h3: ({ children }) => <h3 className="docs-md-h3">{children}</h3>,
  p: ({ children }) => <p className="docs-md-p">{children}</p>,
  ul: ({ children }) => <ul className="docs-md-ul">{children}</ul>,
  ol: ({ children }) => <ol className="docs-md-ol">{children}</ol>,
  li: ({ children }) => <li className="docs-md-li">{children}</li>,
  blockquote: ({ children }) => (
    <blockquote className="docs-md-blockquote">{children}</blockquote>
  ),
  hr: () => <hr className="docs-md-hr" />,
  table: ({ children }) => (
    <div className="docs-table-wrap">
      <table className="docs-table">{children}</table>
    </div>
  ),
  thead: ({ children }) => <thead>{children}</thead>,
  tbody: ({ children }) => <tbody>{children}</tbody>,
  tr: ({ children }) => <tr>{children}</tr>,
  th: ({ children }) => <th>{children}</th>,
  td: ({ children }) => <td>{children}</td>,
  pre: ({ children }) => <pre className="docs-code">{children}</pre>,
  code: ({ className, children, ...props }) => {
    const isBlock = Boolean(className);
    if (isBlock) {
      return (
        <code className={className} {...props}>
          {children}
        </code>
      );
    }
    return (
      <code className="docs-md-inline-code" {...props}>
        {children}
      </code>
    );
  },
  a: ({ href, children }) => {
    const resolved = resolveMarkdownHref(href);
    const isExternal =
      resolved?.startsWith("http") || resolved?.startsWith("mailto:");

    if (isExternal) {
      return (
        <a
          href={resolved}
          className="docs-md-link"
          target="_blank"
          rel="noopener noreferrer"
        >
          {children}
        </a>
      );
    }

    if (resolved?.startsWith("/")) {
      return (
        <Link href={resolved} className="docs-md-link">
          {children}
        </Link>
      );
    }

    return (
      <a href={resolved} className="docs-md-link">
        {children}
      </a>
    );
  },
};

type MarkdownContentProps = {
  content: string;
};

export function MarkdownContent({ content }: MarkdownContentProps) {
  return (
    <article className="docs-markdown w-full max-w-full">
      <ReactMarkdown remarkPlugins={[remarkGfm]} components={components}>
        {content}
      </ReactMarkdown>
    </article>
  );
}
