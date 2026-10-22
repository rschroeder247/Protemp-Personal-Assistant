import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Protemp Personal Assistant | Voice to WhatsApp, Outlook & Planner',
  description:
    'Voice-powered executive assistant for sending WhatsApp messages, creating Outlook tasks & emails, and generating Microsoft Planner tasks.',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className="dark">
      <body className="antialiased selection:bg-indigo-500 selection:text-white">
        {children}
      </body>
    </html>
  );
}
