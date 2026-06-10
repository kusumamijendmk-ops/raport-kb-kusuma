import type {Metadata} from 'next';
import { Inter, Space_Grotesk } from 'next/font/google';
import './globals.css'; // Global styles

const inter = Inter({
  subsets: ['latin'],
  variable: '--font-sans',
  display: 'swap',
});

const spaceGrotesk = Space_Grotesk({
  subsets: ['latin'],
  variable: '--font-display',
  display: 'swap',
});

export const metadata: Metadata = {
  title: 'Aplikasi Raport PAUD Digital',
  description: 'Sistem Pencatatan, Penilaian Perkembangan Anak, dan Cetak Raport Digital PAUD Terintegrasi AI',
};

export default function RootLayout({children}: {children: React.ReactNode}) {
  return (
    <html lang="id" className={`${inter.variable} ${spaceGrotesk.variable}`}>
      <body suppressHydrationWarning className="font-sans antialiased text-gray-800 bg-gray-50/50 min-h-screen selection:bg-indigo-100 selection:text-indigo-800">
        {children}
      </body>
    </html>
  );
}

