import fs from 'fs';
import path from 'path';
import zlib from 'zlib';

// Générateur d'icônes PNG pures avec zlib (sans dépendance externe)
function createPng(width, height, r, g, b) {
  const header = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]); // PNG signature

  // IHDR chunk
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  function makeChunk(type, data) {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length, 0);
    const typeBuf = Buffer.from(type, 'ascii');
    const crcBuf = Buffer.alloc(4);
    const combined = Buffer.concat([typeBuf, data]);
    // CRC calculation
    let crc = 0 ^ (-1);
    for (let i = 0; i < combined.length; i++) {
      crc = (crc >>> 8) ^ table[(crc ^ combined[i]) & 0xff];
    }
    crc = (crc ^ (-1)) >>> 0;
    crcBuf.writeUInt32BE(crc, 0);
    return Buffer.concat([len, combined, crcBuf]);
  }

  // Precompute CRC table
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }

  // Raw scanlines with filter 0
  const scanlines = Buffer.alloc(height * (width * 4 + 1));
  let offset = 0;
  for (let y = 0; y < height; y++) {
    scanlines[offset++] = 0; // filter None
    for (let x = 0; x < width; x++) {
      // Create a nice gradient with an Excel green tone
      const isBorder = (x === 0 || y === 0 || x === width - 1 || y === height - 1);
      const isInner = (x > width * 0.2 && x < width * 0.8 && y > height * 0.2 && y < height * 0.8);
      
      scanlines[offset++] = isInner ? 255 : (isBorder ? 10 : r);   // R
      scanlines[offset++] = isInner ? 255 : (isBorder ? 124 : g);  // G
      scanlines[offset++] = isInner ? 255 : (isBorder ? 65 : b);   // B
      scanlines[offset++] = 255;                                   // Alpha
    }
  }

  const compressedData = zlib.deflateSync(scanlines);
  const ihdrChunk = makeChunk('IHDR', ihdr);
  const idatChunk = makeChunk('IDAT', compressedData);
  const iendChunk = makeChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([header, ihdrChunk, idatChunk, iendChunk]);
}

const assetsDir = path.resolve('public', 'assets');
if (!fs.existsSync(assetsDir)) {
  fs.mkdirSync(assetsDir, { recursive: true });
}

[16, 32, 64, 80, 128].forEach(size => {
  const png = createPng(size, size, 16, 124, 65); // #107C41 Excel Green
  fs.writeFileSync(path.join(assetsDir, `icon-${size}.png`), png);
  console.log(`Generated icon-${size}.png`);
});

// Also create an SVG version
const svgContent = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128" fill="none">
  <rect width="128" height="128" rx="28" fill="#107C41"/>
  <path d="M38 32H74C87.2548 32 98 42.7452 98 56C98 69.2548 87.2548 80 74 80H54V96H38V32Z" fill="white"/>
  <path d="M54 46H74C79.5228 46 84 50.4772 84 56C84 61.5228 79.5228 66 74 66H54V46Z" fill="#107C41"/>
  <circle cx="94" cy="94" r="14" fill="#F59E0B"/>
  <path d="M94 86L96.5 91.5L102 94L96.5 96.5L94 102L91.5 96.5L86 94L91.5 91.5L94 86Z" fill="white"/>
</svg>`;
fs.writeFileSync(path.join(assetsDir, 'icon.svg'), svgContent);
console.log('Generated icon.svg');
