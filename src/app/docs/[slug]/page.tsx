import type { Metadata } from "next";
import { notFound } from "next/navigation";
import {
  CompDocPage,
  getCompDocMetadata,
} from "@/features/docs/comp-doc-page";
import {
  getMarkdownDocBySlug,
  getMarkdownSlugs,
} from "@/features/docs/config/site-map";

type PageProps = {
  params: Promise<{ slug: string }>;
};

export function generateStaticParams() {
  return getMarkdownSlugs().map((slug) => ({ slug }));
}

export async function generateMetadata({
  params,
}: PageProps): Promise<Metadata> {
  const { slug } = await params;
  const { title } = await getCompDocMetadata(slug);
  return { title };
}

export default async function DocsSlugPage({ params }: PageProps) {
  const { slug } = await params;

  if (!getMarkdownDocBySlug(slug)) {
    notFound();
  }

  return <CompDocPage slug={slug} />;
}
