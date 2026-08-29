import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'JRMSU Narrative Report Generator',
  description: 'Create a narrative report and download it as a Word document',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
