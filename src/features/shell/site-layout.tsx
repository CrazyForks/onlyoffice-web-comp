/**
 * 站点根布局：Header + Main + Footer。
 */
import { SiteFooter } from "./site-footer";
import { SiteHeader } from "./site-header";

type SiteLayoutProps = {
  children: React.ReactNode;
};

export function SiteLayout({ children }: SiteLayoutProps) {
  return (
    <div className="flex min-h-dvh flex-col bg-white text-neutral-900">
      <SiteHeader />
      <main className="flex-1">{children}</main>
      <SiteFooter />
    </div>
  );
}
