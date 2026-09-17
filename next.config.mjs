/** @type {import('next').NextConfig} */
const basePath = process.env.BASE_PATH || "";

const nextConfig = {
  output: "export", // static export thuần — chạy trên Vercel hoặc GitHub Pages đều được
  basePath, // GitHub Pages: set BASE_PATH=/tên-repo khi build; Vercel: bỏ trống
  reactStrictMode: true,
};

export default nextConfig;