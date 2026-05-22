import path from "node:path";

/** @type {import('next').NextConfig} */
const nextConfig = {
  output: "standalone",
  outputFileTracingRoot: path.dirname(new URL(import.meta.url).pathname),
  poweredByHeader: false,
};

export default nextConfig;
