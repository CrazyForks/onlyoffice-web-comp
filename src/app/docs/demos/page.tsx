import type { Metadata } from "next";
import { redirect } from "next/navigation";

export const metadata: Metadata = {
  title: "在线示例 — OnlyOffice Web Comp",
};

export default function DocsDemosRoutePage() {
  redirect("/docs/demos/single");
}
