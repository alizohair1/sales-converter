import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'JT Sales Register Converter',
  description: 'Convert Sales Register XLS to 10-Minute Interval Summary',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
