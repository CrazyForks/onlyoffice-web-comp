import type { DocHeading } from "../lib/extract-headings";

type DocTocProps = {
  headings: DocHeading[];
};

export function DocToc({ headings }: DocTocProps) {
  if (headings.length === 0) return null;

  return (
    <aside className="docs-page-toc" aria-label="本页目录">
      <p className="docs-page-toc-label">本页目录</p>
      <nav className="docs-page-toc-nav">
        <ul className="docs-page-toc-list">
          {headings.map((heading) => (
            <li
              key={heading.id}
              className={
                heading.level === 2
                  ? "docs-page-toc-item docs-page-toc-item--nested"
                  : "docs-page-toc-item"
              }
            >
              <a href={`#${heading.id}`} className="docs-page-toc-link">
                {heading.text}
              </a>
            </li>
          ))}
        </ul>
      </nav>
    </aside>
  );
}
