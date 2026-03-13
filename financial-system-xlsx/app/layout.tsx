import './globals.css';
import React from 'react';

export const metadata = {
  title: '财务分析与经营决策系统',
  description: '支持真实 Excel 读写的小型网页系统示例'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
