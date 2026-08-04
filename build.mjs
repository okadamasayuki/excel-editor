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
let safeLib = lib.replace(/<\/script/gi, "<\\/script").replace(/<!--/g, "<\\!--");

// SheetJS のコードページ表には U+FFFD（未定義コードポイントの置き換え文字）が
// そのまま数万個入っている。配信側の検証がこれをデコード失敗とみなすため、
// JS エスケープに置き換える。文字列・正規表現リテラルのどちらでも同じ文字に戻る。
const fffd = (safeLib.match(/�/g) || []).length;
assertAllInStrings(safeLib);
safeLib = safeLib.replace(/�/g, "\\uFFFD");

/**
 * U+FFFD が文字列リテラルの中だけに現れることを確かめる。
 * 中にあるなら "�" は同じコードポイントに戻るので、置換は無害だと保証できる。
 * 将来 SheetJS を差し替えて前提が崩れたらビルドを止める。
 */
function assertAllInStrings(src) {
  let state = 0; // 0:通常 1:' 2:" 3:` 4://  5:/**/ 6:/regex/
  let prev = "";
  const bad = [];
  for (let i = 0; i < src.length; i++) {
    const ch = src[i], nx = src[i + 1];
    if (ch === "�" && state !== 1 && state !== 2 && state !== 3) bad.push({ i, state });
    switch (state) {
      case 0:
        if (ch === "'") state = 1;
        else if (ch === '"') state = 2;
        else if (ch === "`") state = 3;
        else if (ch === "/" && nx === "/") { state = 4; i++; }
        else if (ch === "/" && nx === "*") { state = 5; i++; }
        else if (ch === "/") state = /[)\]}\w$]/.test(prev) ? 0 : 6; // 除算か正規表現か
        break;
      case 1: case 2: case 3: case 6:
        if (ch === "\\") { i++; break; }
        if ((state === 1 && ch === "'") || (state === 2 && ch === '"') ||
            (state === 3 && ch === "`") || (state === 6 && ch === "/")) state = 0;
        break;
      case 4: if (ch === "\n") state = 0; break;
      case 5: if (ch === "*" && nx === "/") { state = 0; i++; } break;
    }
    if (!/\s/.test(ch)) prev = ch;
  }
  if (bad.length) {
    console.error(`U+FFFD が文字列リテラルの外に ${bad.length} 個あります（先頭: 位置 ${bad[0].i}）。`);
    console.error("エスケープすると意味が変わる可能性があるため中止します。");
    process.exit(1);
  }
}

const out = src.replace(marker, () => safeLib);
mkdirSync(join(root, "dist"), { recursive: true });
writeFileSync(join(root, "dist/index.html"), out);

if (/�/.test(out)) {
  console.error("dist に U+FFFD が残っています");
  process.exit(1);
}
const kb = (Buffer.byteLength(out) / 1024).toFixed(0);
console.log(`dist/index.html を出力しました (${kb} KB, U+FFFD を ${fffd} 個エスケープ)`);
