import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: {
    default: "Hệ thống Quản lý Chất lượng QĐ-2429/BYT",
    template: "%s | QLCL 2429/BYT",
  },
  description:
    "Hệ thống quản lý tài liệu, hồ sơ và biểu mẫu chất lượng theo QĐ-2429/BYT — Khoa Vi sinh - Miễn dịch, Bệnh viện Phong - Da liễu TW Quy Hòa.",
  applicationName: "QLCL 2429/BYT",
  authors: [{ name: "Khoa Vi sinh - Miễn dịch, Bệnh viện Phong - Da liễu TW Quy Hòa" }],
  robots: {
    index: false,
    follow: false,
  },
  icons: {
    icon: "/favicon.ico",
  },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  themeColor: "#0E3A41",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html
      lang="vi"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
      suppressHydrationWarning
    >
      <body className="min-h-full flex flex-col bg-[#F4F9F8] text-slate-800">
        <style>{`
          @keyframes appFadeIn {
            from {
              opacity: 0;
              transform: translateY(6px);
            }
            to {
              opacity: 1;
              transform: translateY(0);
            }
          }

          .app-root-transition {
            animation: appFadeIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both;
          }

          @media (prefers-reduced-motion: reduce) {
            .app-root-transition {
              animation: none;
            }
          }
        `}</style>

        <div className="app-root-transition flex flex-col min-h-full">
          {children}
        </div>
      </body>
    </html>
  );
}