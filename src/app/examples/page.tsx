import { redirect } from "next/navigation";

export default function ExamplesRedirectPage() {
  redirect("/docs/demos?tab=multi");
}
