import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Bucephalus Inventory Assistant",
  description: "Excel inventory analysis with optional AI summaries"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <head>
        <script
          dangerouslySetInnerHTML={{
            __html: `window.__officeNativeHistory={pushState:window.history.pushState.bind(window.history),replaceState:window.history.replaceState.bind(window.history)};`
          }}
        />
      </head>
      <body>
        {children}
      </body>
    </html>
  );
}
