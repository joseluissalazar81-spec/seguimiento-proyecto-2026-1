// Generador de íconos PNG para PWA — e-learning UST
const zlib = require('zlib');
const fs   = require('fs');

// ── PNG encoder ──────────────────────────────────────────────────────────────
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
function chunk(type, data) {
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
  ihdr.writeUInt32BE(w,0); ihdr.writeUInt32BE(h,4);
  ihdr[8]=8; ihdr[9]=6; // RGBA
  return Buffer.concat([
    Buffer.from([137,80,78,71,13,10,26,10]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0))
  ]);
}

// ── Dibujado ─────────────────────────────────────────────────────────────────
function createIcon(size) {
  const px = Buffer.alloc(size * size * 4, 0);

  function set(x, y, r, g, b, a=255) {
    x=Math.round(x); y=Math.round(y);
    if (x<0||x>=size||y<0||y>=size) return;
    const i=(y*size+x)*4;
    px[i]=r; px[i+1]=g; px[i+2]=b; px[i+3]=a;
  }
  function rect(x, y, w, h, r, g, b) {
    for (let py=Math.floor(y); py<Math.ceil(y+h); py++)
      for (let px2=Math.floor(x); px2<Math.ceil(x+w); px2++)
        set(px2,py,r,g,b);
  }
  function circle(cx, cy, radius, r, g, b) {
    for (let py=Math.floor(cy-radius); py<=Math.ceil(cy+radius); py++)
      for (let px2=Math.floor(cx-radius); px2<=Math.ceil(cx+radius); px2++)
        if (Math.hypot(px2-cx, py-cy)<=radius) set(px2,py,r,g,b);
  }
  function ring(cx, cy, ro, ri, r, g, b) {
    for (let py=Math.floor(cy-ro-1); py<=Math.ceil(cy+ro+1); py++)
      for (let px2=Math.floor(cx-ro-1); px2<=Math.ceil(cx+ro+1); px2++) {
        const d=Math.hypot(px2-cx, py-cy);
        if (d<=ro && d>=ri && py<=cy) set(px2,py,r,g,b); // upper half arc
      }
  }

  const cr = size * 0.22;

  // Fondo con degradado azul → teal + esquinas redondeadas
  for (let y=0; y<size; y++) {
    for (let x=0; x<size; x++) {
      let inside = true;
      if      (x<cr   && y<cr   && Math.hypot(x-cr,      y-cr      )>cr) inside=false;
      else if (x>size-1-cr && y<cr   && Math.hypot(x-(size-1-cr),y-cr      )>cr) inside=false;
      else if (x<cr   && y>size-1-cr && Math.hypot(x-cr,      y-(size-1-cr))>cr) inside=false;
      else if (x>size-1-cr && y>size-1-cr && Math.hypot(x-(size-1-cr),y-(size-1-cr))>cr) inside=false;

      if (!inside) { set(x,y,255,255,255,0); continue; }

      const t = (x+y)/(2*(size-1));
      set(x, y,
        Math.round(0x3a+(0x1a-0x3a)*t),
        Math.round(0x6d+(0x8c-0x6d)*t),
        Math.round(0xb5+(0x7e-0xb5)*t)
      );
    }
  }

  const S = size/192;

  // Monitor — cuerpo blanco
  rect(32*S, 36*S, 128*S, 90*S, 255,255,255);
  // Pantalla oscura
  rect(40*S, 44*S, 112*S, 74*S, 0x1a,0x3a,0x7a);

  // Gorra de graduación en pantalla (blanca)
  const gcx=96*S, gcy=80*S;
  // Tablero (diamante)
  const dr=22*S;
  for (let y=Math.floor(gcy-dr-1); y<=Math.ceil(gcy+1); y++)
    for (let x=Math.floor(gcx-dr-1); x<=Math.ceil(gcx+dr+1); x++)
      if (Math.abs(x-gcx)/dr + Math.abs(y-gcy)/dr <= 1.0) set(x,y,255,255,255);
  // Ala horizontal
  rect(gcx-30*S, gcy, 60*S, 7*S, 255,255,255);
  // Borla
  rect(gcx+28*S, gcy, 3*S, 18*S, 255,255,255);
  circle(gcx+30*S, gcy+22*S, 5*S, 255,255,255);

  // Soporte monitor
  rect(88*S, 126*S, 16*S, 16*S, 255,255,255);
  // Base
  rect(58*S, 142*S, 76*S, 8*S, 255,255,255);

  // Arcos WiFi debajo del monitor (e-learning online)
  const wCx=96*S, wCy=168*S;
  ring(wCx, wCy, 18*S, 13*S, 255,255,220);
  ring(wCx, wCy, 28*S, 23*S, 255,255,220);
  circle(wCx, wCy, 4*S, 255,255,220);

  return px;
}

// Generar 192×192, 512×512 y 180×180 (Apple)
for (const size of [192, 512, 180]) {
  const name = size===180 ? 'apple-touch-icon.png' : `icon-${size}.png`;
  fs.writeFileSync(name, encodePNG(createIcon(size), size, size));
  console.log('✓', name);
}
