'use client'

import { SiteLayout } from '@/features/shell'

export default function ClientLayoutWrapper({
  children,
}: {
  children: React.ReactNode
}) {
  return <SiteLayout>{children}</SiteLayout>
}
