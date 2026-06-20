import type { Metadata } from "next";
import { TabsMultiPage } from "@/components/onlyoffice-web-demo/tabs-multi-page";

export const metadata: Metadata = {
  title: "多实例 — OnlyOffice MVP",
};

export default function MultiInstancePage() {
  return <TabsMultiPage />;
}
