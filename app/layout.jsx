import "./globals.css";

export const metadata = {
  title: "Excel 自动生成工具",
  description: "Excel smart fill tool with a server-side AI proxy.",
};

export const viewport = {
  width: "device-width",
  initialScale: 1,
};

export default function RootLayout({ children }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
