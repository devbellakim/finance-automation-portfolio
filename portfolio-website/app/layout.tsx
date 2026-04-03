import type { Metadata } from 'next'
import { Inter, JetBrains_Mono } from 'next/font/google'
import './globals.css'

const inter = Inter({
  subsets: ['latin'],
  variable: '--font-inter',
  display: 'swap',
})

const jetbrainsMono = JetBrains_Mono({
  subsets: ['latin'],
  variable: '--font-jetbrains',
  display: 'swap',
})

export const metadata: Metadata = {
  title: 'Finance Automation Engineer | Portfolio',
  description:
    'Automating finance workflows — migrating SAP, Alteryx, VBA to Python & Streamlit. ' +
    'Portfolio of RSU/ESPP reporting, ASC 842 lease automation, Excel → PPT pipelines.',
  keywords: ['finance automation', 'python', 'pandas', 'streamlit', 'SAP', 'alteryx', 'VBA'],
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en" className={`${inter.variable} ${jetbrainsMono.variable}`}>
      <body className="bg-background text-white antialiased">{children}</body>
    </html>
  )
}
