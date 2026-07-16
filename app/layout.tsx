import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "东券禁券监查 · 今日看板",
  description: "每日优惠券禁券监查在线共享看板",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
