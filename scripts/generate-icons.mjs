// Builds the PWA/home-screen icons from the AEON logo at build time, so no
// binary assets have to be committed. Crops just the "AEON" wordmark out of
// the wider "AEON HOME TECH" lockup (rows 45-87 of the 300x152 source,
// found by scanning row brightness — see the deploy conversation), trims
// the residual whitespace, then centers it on a white square with padding.
import sharp from "sharp";
import { mkdirSync, existsSync } from "fs";

mkdirSync("public/icons", { recursive: true });

const LOGO = "public/aeon-logo.jpg";
const MASTER_SIZE = 1024;
const LOGO_WIDTH_RATIO = 0.72; // fraction of the square the wordmark occupies

async function buildMaster() {
  const trimmed = await sharp(LOGO)
    .extract({ left: 0, top: 38, width: 300, height: 56 }) // just the "AEON" line
    .trim({ threshold: 10 })
    .png()
    .toBuffer();
  const meta = await sharp(trimmed).metadata();

  const targetW = Math.round(MASTER_SIZE * LOGO_WIDTH_RATIO);
  const targetH = Math.round(meta.height * (targetW / meta.width));
  const resizedLogo = await sharp(trimmed).resize(targetW, targetH).toBuffer();

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
