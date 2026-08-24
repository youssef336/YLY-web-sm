import type { Metadata, Viewport } from 'next';
import './globals.css';
import { ServiceWorkerRegister } from './service-worker-register';

export const metadata: Metadata = {
  title: 'Social Media — Evaluation Profile',
  description:
    'Social Media: offline-first member evaluation, scoring, strict Excel template injection and WhatsApp reporting.',
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  themeColor: '#060a13',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <head>
        <link rel="manifest" href="/manifest.webmanifest" />
        <link rel="icon" href="/icon.svg" type="image/svg+xml" />
      </head>
      <body>
        {children}
        <ServiceWorkerRegister />
      </body>
    </html>
  );
}