// Builds the PWA/home-screen icons from the AEON logo at build time, so no
// binary assets have to be committed. Crops the full "AEON HOME TECH" lockup
// (not just the wordmark) by scanning the source JPEG's raw pixel brightness
// for a tight bounding box around the artwork, then centers it on a white
// square with padding. (Written by hand instead of sharp's .trim(), which
// threw "bad extract area" against this image in the Linux CI runner.)
import sharp from "sharp";
import { mkdirSync, existsSync } from "fs";

mkdirSync("public/icons", { recursive: true });

const LOGO = "public/aeon-logo.jpg";
const MASTER_SIZE = 1024;
const LOGO_WIDTH_RATIO = 0.72; // fraction of the square the lockup occupies

// Tight bounding box around the non-white artwork — any pixel darker than
// `threshold` (out of 255) counts as part of the logo.
async function findContentBounds(path, threshold = 245) {
  const { data, info } = await sharp(path).greyscale().raw().toBuffer({ resolveWithObject: true });
  const { width, height } = info;
  let minX = width, maxX = -1, minY = height, maxY = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[y * width + x] < threshold) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) throw new Error(`No content found in ${path} below brightness ${threshold} — check the source image or threshold.`);
  return { left: minX, top: minY, width: maxX - minX + 1, height: maxY - minY + 1 };
}

async function buildMaster() {
  const bounds = await findContentBounds(LOGO);
  const cropped = await sharp(LOGO).extract(bounds).png().toBuffer();
  const meta = await sharp(cropped).metadata();

  const targetW = Math.round(MASTER_SIZE * LOGO_WIDTH_RATIO);
  const targetH = Math.round(meta.height * (targetW / meta.width));
  const resizedLogo = await sharp(cropped).resize(targetW, targetH).toBuffer();

  return sharp({
    create: { width: MASTER_SIZE, height: MASTER_SIZE, channels: 4, background: "#ffffff" },
  })
    .composite([{ input: resizedLogo, gravity: "center" }])
    .png()
    .toBuffer();
}

async function withRoundedCorners(buf, size, radiusRatio = 0.19) {
  const radius = Math.round(size * radiusRatio);
  const mask = Buffer.from(
    `<svg width="${size}" height="${size}"><rect width="${size}" height="${size}" rx="${radius}" ry="${radius}" fill="#fff"/></svg>`
  );
  return sharp(buf).resize(size, size).composite([{ input: mask, blend: "dest-in" }]).png().toBuffer();
}

if (existsSync(LOGO)) {
  const master = await buildMaster();

  await withRoundedCorners(master, 192).then((b) => sharp(b).toFile("public/icons/icon-192.png"));
  await withRoundedCorners(master, 512).then((b) => sharp(b).toFile("public/icons/icon-512.png"));
  await sharp(master).resize(180, 180).png().toFile("public/icons/apple-touch-icon.png");
  // Maskable: full-bleed background, no rounding — the OS applies its own mask shape.
  await sharp(master).resize(512, 512).png().toFile("public/icons/icon-512-maskable.png");
  console.log("icons generated from", LOGO);
} else {
  // Fallback placeholder icon so the build never breaks if the logo file is missing.
  for (const [file, size] of [["icon-192.png", 192], ["icon-512.png", 512], ["apple-touch-icon.png", 180], ["icon-512-maskable.png", 512]]) {
    await sharp("public/icon.svg", { density: 384 }).resize(size, size).png().toFile(`public/icons/${file}`);
  }
  console.log("logo not found, generated fallback icons from public/icon.svg");
}
