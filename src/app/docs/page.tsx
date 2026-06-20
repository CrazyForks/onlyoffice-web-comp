import type { Metadata } from "next";
import {
  CompDocPage,
  getCompDocMetadata,
} from "@/features/docs/comp-doc-page";

export async function generateMetadata(): Promise<Metadata> {
  const { title } = await getCompDocMetadata("");
  return { title };
}

export default function DocsOverviewPage() {
  return <CompDocPage slug="" />;
}
