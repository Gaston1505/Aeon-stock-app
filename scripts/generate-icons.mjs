// Rasterizes public/icon*.svg into the PNG sizes PWA manifests and iOS need.
// Runs as part of `npm run build` so no binary assets have to be committed.
import sharp from "sharp";
import { mkdirSync } from "fs";

mkdirSync("public/icons", { recursive: true });

const jobs = [
  { src: "public/icon.svg", out: "public/icons/icon-192.png", size: 192 },
  { src: "public/icon.svg", out: "public/icons/icon-512.png", size: 512 },
  { src: "public/icon.svg", out: "public/icons/apple-touch-icon.png", size: 180 },
  { src: "public/icon-maskable.svg", out: "public/icons/icon-512-maskable.png", size: 512 },
];

for (const job of jobs) {
  await sharp(job.src, { density: 384 }).resize(job.size, job.size).png().toFile(job.out);
  console.log("generated", job.out);
}
