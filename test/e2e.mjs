/**
 * dist/index.html をブラウザで動かして、選択 → 配置 → 生成まで通しで検証する。
 *   node test/e2e.mjs
 * Playwright と SheetJS はグローバル/ローカルどちらでも解決する。
 */
import { readFileSync, writeFileSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import { createServer } from "node:http";

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const require = createRequire(import.meta.url);
const load = (name) => {
  for (const p of [name, `/opt/node22/lib/node_modules/${name}`]) {
    try { return require(p); } catch { /* 次の候補へ */ }
  }
  throw new Error(`${name} が見つかりません`);
};
const { chromium } = load("playwright");
// vendor は CommonJS。package.json が "type": "module" なので明示的に評価して読み込む
const XLSX = (() => {
  const m = { exports: {} };
  new Function("module", "exports", readFileSync(join(root, "vendor/xlsx.full.min.js"), "utf8"))(m, m.exports);
  return m.exports;
})();

const tmp = join(root, ".tmp-test");
rmSync(tmp, { recursive: true, force: true });
mkdirSync(tmp, { recursive: true });

// Artifact 側のラッパ（doctype/head/body）を再現したプレビューを作る
const page_html = readFileSync(join(root, "dist/index.html"), "utf8");
const previewPath = join(tmp, "preview.html");
const wrapped = `<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head><body>${page_html}</body></html>`;
writeFileSync(previewPath, wrapped);

// blob: の download 属性は file:// では無視されるため、実際の配信と同じ http で見る
const server = createServer((req, res) => {
  res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
  res.end(wrapped);
}).listen(0);
await new Promise((r) => server.once("listening", r));
const baseUrl = `http://127.0.0.1:${server.address().port}/`;

const fails = [];
const checks = [];
function check(name, cond, extra) {
  checks.push({ name, ok: !!cond, extra });
  if (!cond) fails.push(name + (extra ? ` — ${extra}` : ""));
  console.log(`${cond ? "  ok  " : " FAIL "} ${name}${extra ? ` (${extra})` : ""}`);
}

// Chromium はロケールに合わせてダウンロード名を正規化するため、C ロケールのままだと
// 日本語ファイル名が "download" に落ちる。実機と同じ UTF-8 で起動する。
const browser = await chromium.launch({
  env: { ...process.env, LANG: "C.utf8", LC_ALL: "C.utf8" },
});
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true });
const page = await ctx.newPage();

const errors = [];
page.on("pageerror", (e) => errors.push(String(e)));
page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

await page.goto(baseUrl);
await page.waitForFunction(() => !!window.__app);

// ---- 1. サンプル読み込み ------------------------------------------------
await page.click("#btnSample");
await page.waitForSelector("#srcGrid:not([hidden])");
const sheetNames = await page.$$eval("#sheetList .sheet-item .nm", (ns) => ns.map((n) => n.textContent));
check("サンプルの3シートが並ぶ", sheetNames.length === 3, sheetNames.join(" / "));
check("先頭シートのセルが描画される",
  (await page.$$eval("#srcGrid .gc", (ns) => ns.filter((n) => n.textContent.trim()).length)) > 10);

check("読み込み後はドロップゾーンが消える", !(await page.isVisible("#dropzone")));
check("ファイル名チップが出る", await page.isVisible("#fileChip"));

// ---- 2. プレビュー上のドラッグで範囲選択 --------------------------------
const cellBox = async (sel) => (await page.locator(sel).first().boundingBox());
const gc = (r, c) => `#srcGrid .gc[data-r="${r}"][data-c="${c}"]`;
const b1 = await cellBox(gc(1, 0));
const b2 = await cellBox(gc(5, 3));
await page.mouse.move(b1.x + 5, b1.y + 5);
await page.mouse.down();
await page.mouse.move(b2.x + 40, b2.y + 12, { steps: 8 });
await page.mouse.up();
let selText = await page.textContent("#selInfo");
check("ドラッグで A2:D6 が選択される", /A2:D6/.test(selText) && /5×4/.test(selText), selText);

// ---- 3. 選択範囲のつまみを出力シートへドラッグ＆ドロップ ----------------
const grab = await page.locator("#srcGrid .selbox .grab").boundingBox();
const dstCell = await page.locator('#dstGrid .gc[data-r="2"][data-c="1"]').boundingBox();
await page.mouse.move(grab.x + grab.width / 2, grab.y + grab.height / 2);
await page.mouse.down();
await page.mouse.move(dstCell.x + 20, dstCell.y + 10, { steps: 12 });
await page.mouse.move(dstCell.x + 22, dstCell.y + 12, { steps: 3 });
await page.mouse.up();
let blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("ドラッグ＆ドロップでブロックが1件できる", blocks.length === 1, blocks.join(","));
let dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("落とした位置 B3 が配置先になる", dest[0] === "抜粋1!B3", dest.join(","));
check("出力グリッドにブロック枠が描かれる", (await page.locator("#dstGrid .blockbox").count()) === 1);
check("出力グリッドに値が流し込まれる（5×4=20セル）",
  (await page.$$eval("#dstGrid .gc", (ns) => ns.filter((n) => n.textContent.trim()).length)) === 20);

// ---- 4. 範囲入力での選択 + 「出力へ配置」クリック ------------------------
await page.fill("#refFrom", "A1");
await page.fill("#refTo", "C3");
await page.click("#btnApplyRef");
selText = await page.textContent("#selInfo");
check("入力欄から A1:C3 を選択できる", /A1:C3/.test(selText), selText);
await page.click("#btnPlace");
const dstCell2 = await page.locator('#dstGrid .gc[data-r="12"][data-c="0"]').boundingBox();
await page.mouse.click(dstCell2.x + 10, dstCell2.y + 10);
blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("クリック配置でブロックが2件になる", blocks.length === 2, blocks.join(" / "));

// ---- 5. 文章コマンド ----------------------------------------------------
await page.fill("#cmdInput", "支店別サマリの2行目から7行目を抜粋1のF3に置く");
await page.press("#cmdInput", "Enter");
blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("文章コマンドで3件目が追加される", blocks.length === 3, blocks.join(" / "));
check("コマンドの範囲解釈が正しい", blocks[2] === "支店別サマリ!A2:C7", blocks[2]);
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("コマンドの配置先が F3", dest[2] === "抜粋1!F3", dest[2]);

// シート追加コマンド
await page.fill("#cmdInput", "シート追加 集計用");
await page.press("#cmdInput", "Enter");
check("シート追加コマンドが効く", (await page.$$eval("#dstTabsHost .tab", (ns) => ns.length)) === 2);

// 新しい出力シートへ、シート一覧からシートまるごとドラッグ
const sheetItem = await page.locator('#sheetList .sheet-item:has-text("商品マスタ")').boundingBox();
const dstCell3 = await page.locator('#dstGrid .gc[data-r="1"][data-c="1"]').boundingBox();
await page.mouse.move(sheetItem.x + 12, sheetItem.y + 10);
await page.mouse.down();
await page.mouse.move(dstCell3.x + 20, dstCell3.y + 10, { steps: 12 });
await page.mouse.move(dstCell3.x + 22, dstCell3.y + 12, { steps: 3 });
await page.mouse.up();
blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("シートまるごとのドラッグで4件目", blocks.length === 4, blocks.join(" / "));
check("シート全体の範囲が取れている", blocks[3] === "商品マスタ!A1:C7", blocks[3]);

// ---- 6. 生成してダウンロード -------------------------------------------
const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 15000 }),
  page.click("#btnGen"),
]);
const outPath = join(tmp, "out.xlsx");
await download.saveAs(outPath);
check("ファイル名が元ファイル由来", download.suggestedFilename() === "サンプル売上_抜粋.xlsx", download.suggestedFilename());

const wb = XLSX.read(readFileSync(outPath), { type: "buffer" });
check("出力ブックのシート構成", JSON.stringify(wb.SheetNames) === JSON.stringify(["抜粋1", "集計用"]), wb.SheetNames.join(","));

const ws = wb.Sheets["抜粋1"];
const at = (ref) => (ws[ref] ? ws[ref].v : undefined);
// ブロック1: 売上明細!A2:D6 → B3（A2は見出し行「日付」）
check("B3 に元の A2（日付）が入る", at("B3") === "日付", String(at("B3")));
check("C3 に元の B2（支店）が入る", at("C3") === "支店", String(at("C3")));
check("E3 に元の D2（数量）が入る", at("E3") === "数量", String(at("E3")));
check("B4 に1件目の日付が入る", typeof at("B4") === "string" && /^2026-/.test(String(at("B4"))), String(at("B4")));
check("E4 が数値として保持される", typeof at("E4") === "number", typeof at("E4"));
// ブロック2: 売上明細!A1:C3 → A13
check("A13 にタイトル「売上明細」が入る", at("A13") === "売上明細", String(at("A13")));
// ブロック3: 支店別サマリ!A2:C7 → F3
check("F3 に「支店」が入る", at("F3") === "支店", String(at("F3")));
check("H4 に東京の売上合計（数値）が入る", typeof at("H4") === "number", String(at("H4")));
// 結合セルの持ち越し（売上明細 A1:E1 の結合は A1:C3 ブロックには収まらないので無し）
const ws2 = wb.Sheets["集計用"];
check("集計用シートに商品マスタが載る", ws2["B2"] && ws2["B2"].v === "商品マスタ", ws2["B2"] && String(ws2["B2"].v));
check("集計用の単価が数値", ws2["C4"] && typeof ws2["C4"].v === "number", ws2["C4"] && String(ws2["C4"].v));

// ---- 7. ブロック操作 ----------------------------------------------------
await page.click("#dstTabsHost .tab >> nth=0");
await page.click("#blockList .block-card >> nth=0");
check("インスペクタが開く", await page.isVisible("#inspAnchor"));
check("インスペクタの空状態ヒントは消える", !(await page.isVisible("#inspEmpty")));
await page.fill("#inspAnchor", "A1");
await page.press("#inspAnchor", "Enter");
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("インスペクタで貼り付け位置を変更できる", dest[0] === "抜粋1!A1", dest[0]);
await page.click("#inspTrans");
blocks = await page.$$eval("#blockList .block-card .row2", (ns) => ns.map((n) => n.textContent));
check("転置が効く", /転置/.test(blocks[0]), blocks[0]);
await page.keyboard.press("Control+z");
await page.keyboard.press("Control+z");
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("Ctrl+Z で戻せる", dest[0] === "抜粋1!B3", dest[0]);

const before = await page.$$eval("#blockList .block-card", (ns) => ns.length);
await page.click("#blockList .block-card >> nth=0 >> .icon-btn");
const after = await page.$$eval("#blockList .block-card", (ns) => ns.length);
check("ブロックを削除できる", after === before - 1, `${before} → ${after}`);

// ---- 8. スクリーンショット ---------------------------------------------
mkdirSync(join(root, "test/shots"), { recursive: true });
await page.screenshot({ path: join(root, "test/shots/light.png") });
const hueLight = await page.$eval("#dstGrid .blockbox", (n) => n.style.getPropertyValue("--bc"));
await page.emulateMedia({ colorScheme: "dark" });
await page.waitForTimeout(250);
const hueDark = await page.$eval("#dstGrid .blockbox", (n) => n.style.getPropertyValue("--bc"));
check("ブロック色がテーマに追随する", hueLight !== hueDark, `${hueLight} → ${hueDark}`);
await page.screenshot({ path: join(root, "test/shots/dark.png") });
await page.evaluate(() => document.documentElement.setAttribute("data-theme", "light"));
await page.waitForTimeout(200);
await page.screenshot({ path: join(root, "test/shots/toggle-light.png") });

// 横スクロールが出ていないか
const overflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
check("本文が横スクロールしない", overflow <= 0, `overflow=${overflow}px`);

// 狭い画面
await page.setViewportSize({ width: 420, height: 860 });
await page.waitForTimeout(250);
const overflowNarrow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
check("狭い画面でも横スクロールしない", overflowNarrow <= 1, `overflow=${overflowNarrow}px`);
await page.screenshot({ path: join(root, "test/shots/narrow.png"), fullPage: false });

check("JSエラーが出ていない", errors.length === 0, errors.slice(0, 3).join(" | "));

await browser.close();
server.close();

console.log(`\n${checks.filter((c) => c.ok).length}/${checks.length} 件成功`);
if (fails.length) { console.error("\n失敗:\n- " + fails.join("\n- ")); process.exit(1); }
console.log("すべて成功");
