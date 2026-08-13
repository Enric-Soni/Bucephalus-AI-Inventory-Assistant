import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  turbopack: { root: process.cwd() },
  headers: async () => [
    {
      source: "/(.*)",
      headers: [
        { key: "Cross-Origin-Opener-Policy", value: "unsafe-none" },
        { key: "Content-Security-Policy", value: "frame-ancestors 'self' https://*.office.com https://*.officeapps.live.com https://*.microsoftonline.com" }
      ]
    }
  ]
};

export default nextConfig;
