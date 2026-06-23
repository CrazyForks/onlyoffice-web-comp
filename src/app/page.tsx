import type { Metadata } from "next";
import { HomePage } from "@/features/marketing";

export const metadata: Metadata = {
  title: "OnlyOffice Web Comp — 浏览器端文档编辑器",
  description:
    "将 OnlyOffice 集成到你的 Web 应用，无需 Document Server，支持 Word / Excel / PPT。",
};

export default function Page() {
  return <HomePage />;
}
