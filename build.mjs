/**
 * src/index.html の `INJECT:XLSX` プレースホルダに SheetJS を埋め込み、
 * 外部リクエストを一切しない単一ファイルを 2 種類つくる。
 *
 *   docs/index.html … GitHub Pages 用の完全な HTML 文書（これが公開されるページ）
 *   dist/index.html … Artifact 用の断片。<html>/<head>/<body> は配信側が付ける
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
if (/�/.test(out)) {
  console.error("出力に U+FFFD が残っています");
  process.exit(1);
}

// --- Artifact 用（断片） ---
mkdirSync(join(root, "dist"), { recursive: true });
writeFileSync(join(root, "dist/index.html"), out);

// --- GitHub Pages 用（完全な HTML 文書） ---
// 断片から <title> と <style> を取り出して <head> に移し、残りを <body> に置く
const title = (out.match(/<title>([\s\S]*?)<\/title>/) || [])[1] || "Excel 抜粋エディタ";
const styles = [...out.matchAll(/<style>[\s\S]*?<\/style>/g)].map((m) => m[0]);
let body = out.replace(/<title>[\s\S]*?<\/title>\s*/, "");
for (const s of styles) body = body.replace(s, "");

const desc = "Excel を読み込み、範囲を選んでドラッグ＆ドロップや文章の指示で配置し、"
  + "抜粋した Excel を書き出すツール。処理はすべてブラウザ内で完結します。";
const favicon = "data:image/svg+xml," + encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
  + '<text y=".95em" font-size="92">\u{1F4D7}</text></svg>');

const pageHtml = `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<title>${title}</title>
<meta name="description" content="${desc}">
<meta property="og:type" content="website">
<meta property="og:title" content="${title}">
<meta property="og:description" content="${desc}">
<link rel="icon" href="${favicon}">
${styles.join("\n")}
</head>
<body>
${body.trim()}
</body>
</html>
`;
mkdirSync(join(root, "docs"), { recursive: true });
writeFileSync(join(root, "docs/index.html"), pageHtml);
writeFileSync(join(root, "docs/.nojekyll"), "");   // Jekyll の処理を通さない

const kb = (n) => (Buffer.byteLength(n) / 1024).toFixed(0);
console.log(`docs/index.html (${kb(pageHtml)} KB) と dist/index.html (${kb(out)} KB) を出力しました`);
console.log(`U+FFFD を ${fffd} 個エスケープ`);
