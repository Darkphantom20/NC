import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'JRMSU Narrative Report Generator',
  description: 'A premium JRMSU narrative report generator with animated sections and DOCX export.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-slate-950 text-slate-100">{children}</body>
    </html>
  );
}
