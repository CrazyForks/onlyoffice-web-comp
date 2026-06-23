import { BrandLogo } from "./brand-logo";
import { SITE_GITHUB } from "./site-ui";

export function SiteFooter() {
  return (
    <footer className="border-t border-neutral-200 bg-neutral-50">
      <div className="mx-auto flex max-w-7xl flex-col gap-4 px-4 py-8 sm:flex-row sm:items-center sm:justify-between sm:px-6">
        <BrandLogo href="/" />
        <p className="text-[13px] text-neutral-500">
          浏览器端 OnlyOffice 编辑器封装 · 无需 Document Server
        </p>
        <a
          href={SITE_GITHUB}
          target="_blank"
          rel="noopener noreferrer"
          className="text-[13px] text-neutral-600 underline-offset-4 hover:text-neutral-950 hover:underline"
        >
          GitHub
        </a>
      </div>
    </footer>
  );
}
