// Resize PNG logo → icon-192.png, icon-512.png, apple-touch-icon.png
const zlib = require('zlib');
const fs   = require('fs');

// ── CRC32 ────────────────────────────────────────────────────────────────────
const crcTbl = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();
function crc32(b, o=0, n=b.length) {
  let c = 0xFFFFFFFF;
  for (let i = o; i < o+n; i++) c = crcTbl[(c^b[i])&0xFF] ^ (c>>>8);
  return (c^0xFFFFFFFF)>>>0;
}

// ── PNG Decode ───────────────────────────────────────────────────────────────
function decodePNG(buf) {
  let pos = 8; // skip signature
  let w, h, bitDepth, colorType;
  const idatBufs = [];

  while (pos < buf.length) {
    const len  = buf.readUInt32BE(pos); pos += 4;
    const type = buf.toString('ascii', pos, pos+4); pos += 4;
    const data = buf.slice(pos, pos+len); pos += len;
    pos += 4; // skip CRC

    if (type === 'IHDR') {
      w = data.readUInt32BE(0); h = data.readUInt32BE(4);
      bitDepth = data[8]; colorType = data[9];
    } else if (type === 'IDAT') {
      idatBufs.push(data);
    } else if (type === 'IEND') break;
  }

  const channels = colorType === 6 ? 4 : colorType === 2 ? 3 : colorType === 4 ? 2 : 1;
  const raw = zlib.inflateSync(Buffer.concat(idatBufs));
  const stride = w * channels;
  const rgba = Buffer.alloc(w * h * 4);

  for (let y = 0; y < h; y++) {
    const filterType = raw[y * (1 + stride)];
    const row  = raw.slice(y * (1 + stride) + 1, y * (1 + stride) + 1 + stride);
    const prev = y > 0 ? raw.slice((y-1)*(1+stride)+1, (y-1)*(1+stride)+1+stride) : Buffer.alloc(stride);

    // Defilter
    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? row[x - channels] : 0;
      const b2 = prev[x];
      const c  = x >= channels ? prev[x - channels] : 0;
      if      (filterType === 1) row[x] = (row[x] + a) & 0xFF;
      else if (filterType === 2) row[x] = (row[x] + b2) & 0xFF;
      else if (filterType === 3) row[x] = (row[x] + Math.floor((a + b2) / 2)) & 0xFF;
      else if (filterType === 4) {
        const p = a + b2 - c;
        const pa = Math.abs(p-a), pb = Math.abs(p-b2), pc = Math.abs(p-c);
        row[x] = (row[x] + (pa<=pb&&pa<=pc?a:pb<=pc?b2:c)) & 0xFF;
      }
    }

    // To RGBA
    for (let x = 0; x < w; x++) {
      const di = (y*w+x)*4;
      if (channels === 4) {
        rgba[di]=row[x*4]; rgba[di+1]=row[x*4+1]; rgba[di+2]=row[x*4+2]; rgba[di+3]=row[x*4+3];
      } else if (channels === 3) {
        rgba[di]=row[x*3]; rgba[di+1]=row[x*3+1]; rgba[di+2]=row[x*3+2]; rgba[di+3]=255;
      } else if (channels === 1) {
        rgba[di]=rgba[di+1]=rgba[di+2]=row[x]; rgba[di+3]=255;
      }
    }
  }
  return { w, h, rgba };
}

// ── PNG Encode ───────────────────────────────────────────────────────────────
function pngChunk(type, data) {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'ascii');
  data.copy(out, 8);
  out.writeUInt32BE(crc32(out, 4, 4+data.length), 8+data.length);
  return out;
}
function encodePNG(rgba, w, h) {
  const raw = Buffer.alloc(h * (1 + w * 4));
  for (let y = 0; y < h; y++) {
    raw[y*(1+w*4)] = 0;
    rgba.copy(raw, y*(1+w*4)+1, y*w*4, (y+1)*w*4);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w,0); ihdr.writeUInt32BE(h,4); ihdr[8]=8; ihdr[9]=6;
  return Buffer.concat([
    Buffer.from([137,80,78,71,13,10,26,10]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0))
  ]);
}

// ── Bilinear resize ──────────────────────────────────────────────────────────
function resize(src, sw, sh, dw, dh) {
  const dst = Buffer.alloc(dw * dh * 4);
  const xr = sw / dw, yr = sh / dh;
  for (let dy = 0; dy < dh; dy++) {
    for (let dx = 0; dx < dw; dx++) {
      const sx = (dx + 0.5) * xr - 0.5;
      const sy = (dy + 0.5) * yr - 0.5;
      const x0 = Math.max(0, Math.floor(sx)), x1 = Math.min(sw-1, x0+1);
      const y0 = Math.max(0, Math.floor(sy)), y1 = Math.min(sh-1, y0+1);
      const wx = sx - x0, wy = sy - y0;
      const di = (dy*dw+dx)*4;
      for (let c = 0; c < 4; c++) {
        const a00 = src[(y0*sw+x0)*4+c];
        const a10 = src[(y0*sw+x1)*4+c];
        const a01 = src[(y1*sw+x0)*4+c];
        const a11 = src[(y1*sw+x1)*4+c];
        dst[di+c] = Math.round(a00*(1-wx)*(1-wy)+a10*wx*(1-wy)+a01*(1-wx)*wy+a11*wx*wy);
      }
    }
  }
  return dst;
}

// ── Fit logo centered on square canvas ──────────────────────────────────────
function fitOnSquare(srcRgba, sw, sh, size, bgR=0x3a, bgG=0x6d, bgB=0xb5) {
  const pad = Math.round(size * 0.08);
  const area = size - pad * 2;
  const scale = Math.min(area / sw, area / sh);
  const dw = Math.round(sw * scale);
  const dh = Math.round(sh * scale);
  const ox = Math.round((size - dw) / 2);
  const oy = Math.round((size - dh) / 2);

  const scaled = resize(srcRgba, sw, sh, dw, dh);
  const canvas = Buffer.alloc(size * size * 4);

  // Fondo con degradado azul→teal
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      const t = (x + y) / (2 * (size - 1));
      const i = (y * size + x) * 4;
      canvas[i]   = Math.round(bgR + (0x1a - bgR) * t);
      canvas[i+1] = Math.round(bgG + (0x8c - bgG) * t);
      canvas[i+2] = Math.round(bgB + (0x7e - bgB) * t);
      canvas[i+3] = 255;
    }
  }

  // Pegar logo centrado (alpha blending)
  for (let y = 0; y < dh; y++) {
    for (let x = 0; x < dw; x++) {
      const cx = ox + x, cy = oy + y;
      if (cx < 0 || cx >= size || cy < 0 || cy >= size) continue;
      const si = (y * dw + x) * 4;
      const di = (cy * size + cx) * 4;
      const a  = scaled[si+3] / 255;
      canvas[di]   = Math.round(scaled[si]   * a + canvas[di]   * (1-a));
      canvas[di+1] = Math.round(scaled[si+1] * a + canvas[di+1] * (1-a));
      canvas[di+2] = Math.round(scaled[si+2] * a + canvas[di+2] * (1-a));
      canvas[di+3] = 255;
    }
  }
  return canvas;
}

// ── Run ──────────────────────────────────────────────────────────────────────
const { w, h, rgba } = decodePNG(fs.readFileSync('logo-original.png'));
console.log(`Original: ${w}×${h}`);

for (const [size, name] of [[192,'icon-192.png'],[512,'icon-512.png'],[180,'apple-touch-icon.png']]) {
  const icon = fitOnSquare(rgba, w, h, size);
  fs.writeFileSync(name, encodePNG(icon, size, size));
  console.log(`✓ ${name} (${size}×${size})`);
}
