/**
 * 顶栏导航配置：路由、文案与激活态判断（文档 / 示例路径互不干扰）。
 */
export type SiteNavItem = {
  href: string;
  label: string;
  isActive: (pathname: string) => boolean;
};

export const SITE_NAV: SiteNavItem[] = [
  {
    href: "/",
    label: "主页",
    isActive: (pathname) => pathname === "/",
  },
  {
    href: "/docs",
    label: "文档",
    isActive: (pathname) =>
      pathname === "/docs" ||
      (pathname.startsWith("/docs/") && !pathname.startsWith("/docs/demos")),
  },
  {
    href: "/docs/demos/single",
    label: "示例",
    isActive: (pathname) => pathname.startsWith("/docs/demos"),
  },
];

export function isDemoPath(pathname: string) {
  return pathname.startsWith("/docs/demos");
}
