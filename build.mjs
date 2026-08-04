/**
 * src/index.html の `INJECT:XLSX` プレースホルダに SheetJS を埋め込み、
 * 外部リクエストを一切しない単一ファイル dist/index.html を作る。
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const src = readFileSync(join(root, "src/index.html"), "utf8");
const lib = readFileSync(join(root, "vendor/xlsx.full.min.js"), "utf8");

const marker = "/* INJECT:XLSX */";
if (!src.includes(marker)) {
  console.error("プレースホルダ " + marker + " が src/index.html にありません");
  process.exit(1);
}

// インラインスクリプト内で HTML パーサを壊さないようエスケープする
const safeLib = lib.replace(/<\/script/gi, "<\\/script").replace(/<!--/g, "<\\!--");

const out = src.replace(marker, () => safeLib);
mkdirSync(join(root, "dist"), { recursive: true });
writeFileSync(join(root, "dist/index.html"), out);

const kb = (Buffer.byteLength(out) / 1024).toFixed(0);
console.log(`dist/index.html を出力しました (${kb} KB)`);
