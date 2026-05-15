import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: 'export',
  distDir: 'www',
  images: { unoptimized: true },
};

export default nextConfig;
