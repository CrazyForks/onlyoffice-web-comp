import type { Metadata } from "next";
import ClientLayoutWrapper from './ClientLayoutWrapper';
import "./globals.css";

export const metadata: Metadata = {
  title: "OnlyOffice Web Comp",
  description: "浏览器端 OnlyOffice 编辑器组件，无需 Document Server",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body className="antialiased">
        <ClientLayoutWrapper>
          {children}
        </ClientLayoutWrapper>
      </body>
    </html>
  );
}
