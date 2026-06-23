import type { Metadata } from "next";
import { Inter } from "next/font/google";
import ClientLayoutWrapper from './ClientLayoutWrapper';
import "./globals.css";

const inter = Inter({
  subsets: ["latin"],
  display: "swap",
  variable: "--font-inter",
});

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
    <html lang="zh-CN" className={inter.variable}>
      <body className={`${inter.className} antialiased`}>
        <ClientLayoutWrapper>
          {children}
        </ClientLayoutWrapper>
      </body>
    </html>
  );
}
