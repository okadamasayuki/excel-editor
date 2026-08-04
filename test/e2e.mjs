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

// 実際に GitHub Pages で配信されるファイルそのものを検証する
const pageHtml = readFileSync(join(root, "docs/index.html"), "utf8");
// Artifact 用の断片は、配信側のラッパ（doctype/head/body）を再現して煙テストする
const fragment = readFileSync(join(root, "dist/index.html"), "utf8");
const wrappedFragment = `<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head><body>${fragment}</body></html>`;
writeFileSync(join(tmp, "artifact-preview.html"), wrappedFragment);

// blob: の download 属性は file:// では無視されるため、実際の配信と同じ http で見る
const server = createServer((req, res) => {
  res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
  res.end(req.url.startsWith("/artifact") ? wrappedFragment : pageHtml);
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

// 社内ファイルを扱う道具なので「外に出ていない」ことを実測で押さえる。
// ページ本体の取得より後に発生した通信をすべて記録する。
const requests = [];
page.on("request", (r) => requests.push({ url: r.url(), method: r.method(), type: r.resourceType() }));

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

// ---- 9. 手順書（同じ処理の使い回し / 引き継ぎ） -------------------------
page.on("dialog", (d) => d.accept());   // 上書き確認などは通す
await page.click("#tabScript");
check("手順書タブに切り替わる", await page.isVisible("#scText"));

// 今の配置から手順書を自動生成できる
await page.click("#scFromBlocks");
const generated = await page.inputValue("#scText");
check("今の配置から手順書ができる",
  /元ファイル: サンプル売上\.xlsx/.test(generated) && /シート追加 抜粋1/.test(generated)
  && /を抜粋1の[A-Z]+\d+に置く/.test(generated),
  generated.split("\n").filter((l) => l.trim()).slice(0, 3).join(" / "));

// 繰り返し + 続けて置く。10/12/17/20 行目を 3 シートぶん縦に積む
const recipe = `# 元ファイル: サンプル売上.xlsx
説明: 各行を3シートぶん並べる

シート追加 行そろえ
繰り返し 行 = 10, 12, 17, 20
  売上明細の{行}行目を行そろえのA1に置く
  支店別サマリの{行}行目を続けて置く
  商品マスタの{行}行目を続けて置く
ここまで`;
await page.fill("#scText", recipe);
await page.click("#scRun");
await page.waitForTimeout(200);

const rBlocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("繰り返しが 4値 × 3行 = 12件に展開される", rBlocks.length === 12, `${rBlocks.length}件`);
const rDest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("1件目は A1", rDest[0] === "行そろえ!A1", rDest[0]);
check("続けて置くで1行ずつ下に積まれる", rDest[1] === "行そろえ!A2" && rDest[2] === "行そろえ!A3",
  rDest.slice(0, 3).join(" / "));
check("2周目も続けて積まれる", rDest[3] === "行そろえ!A4" && rDest[11] === "行そろえ!A12",
  rDest.slice(3, 4) + " … " + rDest[11]);
check("{行} が値に置き換わる", rBlocks[0] === "売上明細!A10:E10" && rBlocks[4] === "支店別サマリ!A12:C12",
  rBlocks[0] + " / " + rBlocks[4]);

// 出力の中身を確認（12行目 = 4周目の1行目 = 売上明細の20行目）
const [dl2] = await Promise.all([page.waitForEvent("download"), page.click("#btnGen")]);
const out2 = join(tmp, "recipe.xlsx");
await dl2.saveAs(out2);
const wb2 = XLSX.read(readFileSync(out2), { type: "buffer" });
const ws3 = wb2.Sheets["行そろえ"];
const src1 = XLSX.read(readFileSync(out2), { type: "buffer" }); // 参照用
check("手順書の出力シート名が反映される", !!ws3, wb2.SheetNames.join(","));
check("1行目に売上明細の10行目が入る", ws3 && ws3["B1"] && typeof ws3["B1"].v === "string", ws3 && ws3["B1"] && String(ws3["B1"].v));
check("10行分が縦に並ぶ", ws3 && ws3["!ref"] === "A1:E12", ws3 && ws3["!ref"]);

// 「下に続ける」と「右に続ける」を混ぜても、周の先頭は左端に戻る（階段状にならない）
await page.fill("#scText", `シート追加 横並び
売上明細のA2からE2を横並びのA1に置く
繰り返し 行 = 10, 12, 17
  売上明細の{行}行目を続けて置く
  支店別サマリの3行目を右に続けて置く
ここまで`);
await page.click("#scRun");
await page.waitForTimeout(200);
const zig = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("下に続けると行の左端に戻る",
  zig.join(",") === "横並び!A1,横並び!A2,横並び!F2,横並び!A3,横並び!F3,横並び!A4,横並び!F4",
  zig.join(","));

// 薄いブロックを積んでも中身が読めるよう、ラベルは選択中/ホバー中だけ
const tagsShown = await page.$$eval("#dstGrid .blockbox .tag",
  (ns) => ns.filter((n) => getComputedStyle(n).visibility === "visible").length);
check("ブロックのラベルは既定で隠れている", tagsShown <= 1, `${tagsShown}件が表示中`);
await page.hover("#dstGrid .blockbox >> nth=0");
check("ホバーでラベルが出る",
  await page.$eval("#dstGrid .blockbox >> nth=0 >> .tag", (n) => getComputedStyle(n).visibility === "visible"));

// 「1回ぶん配置 → 今の配置から作る → 繰り返しにする」の流れ
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(`シート追加 まとめ
売上明細の10行目をまとめのA1に置く
支店別サマリの10行目を右に続けて置く`, true);
});
await page.click("#scFromBlocks");
await page.click("#scRepeat");
const wrapped = await page.inputValue("#scText");
check("繰り返しにするで包まれる",
  /繰り返し 行 = 10\n {2}売上明細のA\{行\}:E\{行\}をまとめのA1に置く\n {2}支店別サマリのA\{行\}:C\{行\}をまとめのF1に置く\nここまで/.test(wrapped),
  wrapped.split("\n").filter((l) => /繰り返し|\{行\}|ここまで/.test(l)).join(" / "));
check("シート追加は繰り返しの外に残る",
  /シート追加 まとめ\n繰り返し/.test(wrapped),
  wrapped.split("\n").filter((l) => /シート追加/.test(l)).join(","));
check("値の欄が選択された状態になる",
  await page.evaluate(() => {
    const ta = document.getElementById("scText");
    return ta.value.slice(ta.selectionStart, ta.selectionEnd);
  }) === "10");

// 選択部分に値を足してそのまま実行できる
await page.evaluate(() => {
  const ta = document.getElementById("scText");
  ta.value = ta.value.replace("繰り返し 行 = 10", "繰り返し 行 = 10, 12, 17, 20");
});
await page.click("#scRun");
await page.waitForTimeout(200);
const wrapDest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("包んだ手順書がそのまま4周ぶん動く",
  wrapDest.join(",") === "まとめ!A1,まとめ!F1,まとめ!A2,まとめ!F2,まとめ!A3,まとめ!F3,まとめ!A4,まとめ!F4",
  wrapDest.join(","));

// 数字を含むシート名を行番号として巻き込まない
const guarded = await page.evaluate(() => {
  const A = window.__app;
  A.S.sheets.push({ name: "支店10", ws: {}, rows: 1, cols: 1, usedRows: 1, usedCols: 1, merges: [], colsMeta: [] });
  const r = A.commonRowNumber(["支店10のA10:E10"]);
  A.S.sheets.pop();
  return r;
});
check("シート名の数字も候補には入る（保護は置換時）", guarded === "10", String(guarded));

// 保存 → 別の手順書に差し替え → 読み戻し
await page.fill("#scName", "月次テスト");
await page.click("#scSave");
await page.fill("#scText", "# 消してよい内容");
await page.selectOption("#scList", "月次テスト");
check("保存した手順書を読み戻せる", (await page.inputValue("#scText")).includes("繰り返し 行 = 10, 12, 17"));
check("保存件数が表示される", /保存済み 1件/.test(await page.textContent("#recipeState")));

// 別ブラウザの人に渡す想定でファイルに書き出す
const onScreen = await page.inputValue("#scText");
const [dl3] = await Promise.all([page.waitForEvent("download"), page.click("#scExport")]);
check("手順書をファイルに書き出せる", dl3.suggestedFilename() === "月次テスト.txt", dl3.suggestedFilename());
const exported = join(tmp, "recipe.txt");
await dl3.saveAs(exported);
const exportedText = readFileSync(exported, "utf8");
check("書き出した内容が画面の手順書と一致する", exportedText === onScreen,
  `${exportedText.length}文字 / ${onScreen.length}文字`);
check("書き出した手順書がそのまま読める文章",
  /繰り返し .+ = /.test(exportedText) && /に置く/.test(exportedText),
  exportedText.split("\n")[0]);

// 存在しないシートは黙って別シートに逃げず、行番号つきで止まる
await page.fill("#scText", "存在しないシートのA1:B2を抜粋1のA1に置く");
await page.click("#scRun");
await page.waitForTimeout(150);
check("無いシートは行番号つきで報告される",
  /1行目.*見つかりません/.test(await page.textContent("#log")), (await page.textContent("#log")).slice(0, 60));

// ---- 9b. 読み込んだExcelがブラウザに保存されないことの実測 ---------------
const storage = await page.evaluate(async () => {
  const ls = {};
  for (let i = 0; i < localStorage.length; i++) { const k = localStorage.key(i); ls[k] = localStorage.getItem(k); }
  const ss = {};
  for (let i = 0; i < sessionStorage.length; i++) { const k = sessionStorage.key(i); ss[k] = sessionStorage.getItem(k); }
  let dbs = [];
  try { dbs = (await indexedDB.databases()).map((d) => d.name); } catch (e) { dbs = []; }
  return { ls, ss, cookie: document.cookie, dbs };
});
check("localStorage に置くのは手順書だけ",
  Object.keys(storage.ls).every((k) => k === "excel-extract-recipes-v1"),
  Object.keys(storage.ls).join(","));
check("sessionStorage / Cookie / IndexedDB は未使用",
  Object.keys(storage.ss).length === 0 && storage.cookie === "" && storage.dbs.length === 0,
  `ss=${Object.keys(storage.ss).length} cookie="${storage.cookie}" idb=${storage.dbs.join(",")}`);
// セルの中身（サンプルの実データ）が保存領域に残っていないこと
const leaked = ["名古屋", "デスク", "32000", "111600"].filter((w) => JSON.stringify(storage).includes(w));
check("セルの値は保存領域に残らない", leaked.length === 0, leaked.join(","));
// 再読み込みで読み込み済みブックが消えること
await page.reload();
await page.waitForFunction(() => !!window.__app);
check("再読み込みで読み込んだブックは消える",
  await page.evaluate(() => window.__app.S.wb === null && window.__app.S.sheets.length === 0));
check("手順書は再読み込み後も残る",
  (await page.evaluate(() => localStorage.getItem("excel-extract-recipes-v1"))) !== null);

// ---- 10. 外部に一切送信していないことの実測 -----------------------------
const external = requests.filter((r) => !r.url.startsWith(baseUrl) && !r.url.startsWith("data:") && !r.url.startsWith("blob:"));
check("ページ取得以外の外部通信が 0 件", external.length === 0,
  external.slice(0, 3).map((r) => `${r.method} ${r.url}`).join(" | "));
const nonDoc = requests.filter((r) => r.resourceType !== "document" && r.type !== "document");
check("画像・スクリプト等の追加取得も 0 件", nonDoc.length === 0,
  nonDoc.slice(0, 3).map((r) => `${r.type} ${r.url}`).join(" | "));

// ---- 9. Artifact 用の断片も動くか（煙テスト） ---------------------------
const p2 = await ctx.newPage();
const errors2 = [];
p2.on("pageerror", (e) => errors2.push(String(e)));
await p2.goto(baseUrl + "artifact");
await p2.waitForFunction(() => !!window.__app, { timeout: 10000 });
await p2.click("#btnSample");
await p2.waitForSelector("#srcGrid:not([hidden])");
check("Artifact 断片版も起動する",
  (await p2.$$eval("#sheetList .sheet-item", (ns) => ns.length)) === 3 && errors2.length === 0,
  errors2.slice(0, 2).join(" | "));

await browser.close();
server.close();

console.log(`\n${checks.filter((c) => c.ok).length}/${checks.length} 件成功`);
if (fails.length) { console.error("\n失敗:\n- " + fails.join("\n- ")); process.exit(1); }
console.log("すべて成功");
