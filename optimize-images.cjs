// One-shot script to optimize hero and focus area images for web.
// Reads PNGs from public/images/, writes JPGs at sane web dimensions.
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const IMAGES_DIR = path.join(__dirname, 'public', 'images');

// What to optimize and target dimensions
const TARGETS = [
  { src: 'hero-atlas.png',           out: 'hero-atlas.jpg',           width: 1600, quality: 82 },
  { src: 'focus-1-neighborhood.png', out: 'focus-1-neighborhood.jpg', width: 1200, quality: 82 },
  { src: 'focus-2-toolkit.png',      out: 'focus-2-toolkit.jpg',      width: 1200, quality: 82 },
  { src: 'focus-3-resolution.png',   out: 'focus-3-resolution.jpg',   width: 1200, quality: 82 },
  { src: 'michael-wright.jpg',       out: 'michael-wright-opt.jpg',   width: 900,  quality: 85 },
];

async function run() {
  let totalIn = 0;
  let totalOut = 0;
  for (const t of TARGETS) {
    const inPath = path.join(IMAGES_DIR, t.src);
    const outPath = path.join(IMAGES_DIR, t.out);
    if (!fs.existsSync(inPath)) {
      console.log(`SKIP (missing): ${t.src}`);
      continue;
    }
    const inSize = fs.statSync(inPath).size;
    await sharp(inPath)
      .resize({ width: t.width, withoutEnlargement: true })
      .jpeg({ quality: t.quality, mozjpeg: true })
      .toFile(outPath);
    const outSize = fs.statSync(outPath).size;
    totalIn += inSize;
    totalOut += outSize;
    const pct = (100 * (1 - outSize / inSize)).toFixed(1);
    console.log(
      `${t.src.padEnd(32)} ${(inSize/1024/1024).toFixed(2)} MB  ->  ${t.out.padEnd(32)} ${(outSize/1024).toFixed(0)} KB   (-${pct}%)`
    );
  }
  console.log('-'.repeat(60));
  console.log(
    `TOTAL:                            ${(totalIn/1024/1024).toFixed(2)} MB  ->                                ${(totalOut/1024/1024).toFixed(2)} MB   (-${(100*(1-totalOut/totalIn)).toFixed(1)}%)`
  );
}

run().catch(e => { console.error(e); process.exit(1); });
