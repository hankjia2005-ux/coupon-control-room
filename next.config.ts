import type { NextConfig } from "next";

const isGithubPages = process.env.GITHUB_PAGES === "true";
const repositoryName = process.env.GITHUB_REPOSITORY?.split("/")[1] ?? "";
const isUserSite = repositoryName.endsWith(".github.io");

const nextConfig: NextConfig = {
  output: "export",
  basePath: isGithubPages && repositoryName && !isUserSite ? `/${repositoryName}` : "",
  images: { unoptimized: true },
};

export default nextConfig;
