import { redirect } from "next/navigation";

export default function MultiRedirectPage() {
  redirect("/docs/demos?tab=multi");
}
