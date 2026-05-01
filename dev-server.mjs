import http from "node:http";
import https from "node:https";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.PORT || 3000);
const certPath = path.join(__dirname, "certs", "localhost.pem");
const keyPath = path.join(__dirname, "certs", "localhost-key.pem");

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".mjs", "text/javascript; charset=utf-8"],
  [".xml", "application/xml; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"]
]);

const crcTable = new Uint32Array(256).map((_, index) => {
  let c = index;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return c >>> 0;
});

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const typeBuffer = Buffer.from(type);
  const length = Buffer.alloc(4);
  const crc = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])));
  return Buffer.concat([length, typeBuffer, data, crc]);
}

function createIcon(size) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  const teal = [21, 96, 100, 255];
  const mint = [193, 255, 214, 255];
  const white = [255, 255, 255, 255];
  const center = (size - 1) / 2;
  const radius = size * 0.42;

  for (let y = 0; y < size; y += 1) {
    const row = y * (size * 4 + 1);
    raw[row] = 0;

    for (let x = 0; x < size; x += 1) {
      const offset = row + 1 + x * 4;
      const distance = Math.hypot(x - center, y - center);
      const base = distance < radius ? mint : teal;
      raw[offset] = base[0];
      raw[offset + 1] = base[1];
      raw[offset + 2] = base[2];
      raw[offset + 3] = base[3];

      const plate = Math.hypot(x - center, y - center) < size * 0.2;
      const forkHandle = x > size * 0.25 && x < size * 0.31 && y > size * 0.23 && y < size * 0.73;
      const forkHead = x > size * 0.2 && x < size * 0.36 && y > size * 0.23 && y < size * 0.36;
      const spoon = Math.hypot(x - size * 0.68, y - size * 0.33) < size * 0.07 ||
        (x > size * 0.66 && x < size * 0.72 && y > size * 0.38 && y < size * 0.75);

      if (plate || forkHandle || forkHead || spoon) {
        raw[offset] = white[0];
        raw[offset + 1] = white[1];
        raw[offset + 2] = white[2];
        raw[offset + 3] = white[3];
      }
    }
  }

  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  return Buffer.concat([
    signature,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", zlib.deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0))
  ]);
}

async function serve(request, response) {
  const url = new URL(request.url, `http://${request.headers.host}`);
  let pathname = decodeURIComponent(url.pathname);

  if (pathname === "/") {
    pathname = "taskpane.html";
  }

  const iconMatch = pathname.match(/^\/assets\/icon-(16|32|64|80|128)\.png$/);
  if (iconMatch) {
    const size = Number(iconMatch[1]);
    response.writeHead(200, {
      "Content-Type": "image/png",
      "Cache-Control": "no-store"
    });
    response.end(createIcon(size));
    return;
  }

  const filePath = path.normalize(path.join(__dirname, pathname));
  if (!filePath.startsWith(__dirname)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  try {
    const content = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type": mimeTypes.get(path.extname(filePath)) || "application/octet-stream",
      "Cache-Control": "no-store"
    });
    response.end(content);
  } catch {
    response.writeHead(404);
    response.end("Not found");
  }
}

async function main() {
  const hasCert = existsSync(certPath) && existsSync(keyPath);
  const server = hasCert
    ? https.createServer({ cert: await readFile(certPath), key: await readFile(keyPath) }, serve)
    : http.createServer(serve);

  server.listen(port, () => {
    const protocol = hasCert ? "https" : "http";
    console.log(`Frokost add-in running at ${protocol}://localhost:${port}`);
    if (!hasCert) {
      console.log("Outlook sideloading normally requires HTTPS. Add certs/localhost.pem and certs/localhost-key.pem for HTTPS.");
    }
  });
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
