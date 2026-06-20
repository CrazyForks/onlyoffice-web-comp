import { redirect } from "next/navigation";

export default function DocsDemosSingleRedirectPage() {
  redirect("/docs/demos?tab=single");
}
