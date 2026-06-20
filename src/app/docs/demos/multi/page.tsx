import { redirect } from "next/navigation";

export default function DocsDemosMultiRedirectPage() {
  redirect("/docs/demos?tab=multi");
}
