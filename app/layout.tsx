import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "RECAP Filler — POG Category Auto-Fill",
  description: "เติมข้อมูล DIVISION / DEPT / SUB-DEPT / Class / PLANOGRAM อัตโนมัติจากไฟล์ 100 ช่อง",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="th">
      <body>{children}</body>
    </html>
  );
}
