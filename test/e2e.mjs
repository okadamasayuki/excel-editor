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

/** 範囲を選ぶ（画面の入力欄は廃止したので、アプリの API で指定する） */
async function selectRange(from, to) {
  await page.evaluate(([f, t]) => {
    const A = window.__app, a = A.parseCell(f), b = A.parseCell(t);
    A.setSelection(a.r, a.c, b.r, b.c);
    A.revealSelection();
  }, [from, to]);
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

// ---- 3b. 端まで引っ張ると自動スクロールして隠れた行・列が出る ------------
// 選択の中から始めると「出力へ運ぶドラッグ」になるので、選択の外から始める
const gridBox = await page.locator("#srcGrid").boundingBox();
const startCell = await cellBox(gc(9, 0));
await page.mouse.move(startCell.x + 5, startCell.y + 5);
await page.mouse.down();
// 右端で保持 → 見えていない列の方向へ送られ続ける
await page.mouse.move(gridBox.x + gridBox.width - 6, startCell.y + 5, { steps: 6 });
await page.waitForTimeout(450);
const scrolledX = await page.evaluate(() => document.getElementById("srcGrid").scrollLeft);
const selAfterX = await page.evaluate(() => window.__app.S.sel.c2);
check("右端で引っ張ると横に自動スクロールする", scrolledX > 100, `scrollLeft=${Math.round(scrolledX)}`);
check("見えていなかった列まで選択が伸びる", selAfterX >= 8, `c2=${selAfterX}`);
// 下端で保持 → 行方向にも送られる
await page.mouse.move(gridBox.x + gridBox.width / 2, gridBox.y + gridBox.height - 6, { steps: 6 });
await page.waitForTimeout(450);
const scrolledY = await page.evaluate(() => document.getElementById("srcGrid").scrollTop);
const selAfterY = await page.evaluate(() => window.__app.S.sel.r2);
check("下端で引っ張ると縦に自動スクロールする", scrolledY > 100, `scrollTop=${Math.round(scrolledY)}`);
check("見えていなかった行まで選択が伸びる", selAfterY >= 20, `r2=${selAfterY}`);
await page.mouse.up();
// 離したら止まる
const restA = await page.evaluate(() => document.getElementById("srcGrid").scrollTop);
await page.waitForTimeout(300);
const restB = await page.evaluate(() => document.getElementById("srcGrid").scrollTop);
check("離したら自動スクロールが止まる", restA === restB, `${Math.round(restA)} → ${Math.round(restB)}`);

// 横スクロール中でも、貼り付いている行見出しは見出しとして扱われること
// （描画は画面外にも少し余分に作るので、位置は座標で直接指定する）
const gb = await page.locator("#srcGrid").boundingBox();
await page.mouse.click(gb.x + 20, gb.y + 26 + 40);
const selAfterHead = await page.evaluate(() => window.__app.S.sel);
check("横スクロール中でも行見出しで行全体を選べる",
  selAfterHead.c1 === 0 && selAfterHead.r1 === selAfterHead.r2,
  JSON.stringify(selAfterHead));
await page.keyboard.press("Escape");
await page.evaluate(() => { const w = document.getElementById("srcGrid"); w.scrollTop = 0; w.scrollLeft = 0; });
await page.waitForTimeout(80);

// ---- 4. 範囲を指定しての選択 + 「出力へ配置」クリック --------------------
await selectRange("A1", "C3");
selText = await page.textContent("#selInfo");
check("A1:C3 を選択できる", /A1:C3/.test(selText), selText);
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
// （直前のコマンド配置で出力側がその位置まで動いているので、先頭に戻してから掴む）
await page.evaluate(() => {
  const w = document.getElementById("dstGrid");
  w.scrollTop = 0; w.scrollLeft = 0;
});
await page.waitForTimeout(80);
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
// 狭い画面でも、グリッドは自分の中でスクロールすること
// （高さを持たないと行数ぶん伸びて、ページが際限なく長くなる）
const narrowGrid = await page.evaluate(() => {
  const w = document.getElementById("srcGrid");
  return { clientH: w.clientHeight, scrollH: w.scrollHeight };
});
check("狭い画面でもグリッドの高さが抑えられる",
  narrowGrid.clientH < 700 && narrowGrid.clientH < narrowGrid.scrollH,
  `clientH=${narrowGrid.clientH} scrollH=${narrowGrid.scrollH}`);
await page.screenshot({ path: join(root, "test/shots/narrow.png"), fullPage: false });
await page.setViewportSize({ width: 1440, height: 900 });   // 以降は通常の画面で検証する
await page.waitForTimeout(150);

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

// ---- 9c. 読み込み中の表示 ------------------------------------------------
// 大きめのブックを作って、読み込み中に状態が出ることを確かめる
const bigRows = [["日付", "支店", "商品", "数量", "金額"]];
for (let i = 0; i < 40000; i++) bigRows.push([`2026-01-${(i % 28) + 1}`, `支店${i % 20}`, `商品${i % 50}`, i % 9, i * 137]);
const bigWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(bigWb, XLSX.utils.aoa_to_sheet(bigRows), "大量データ");
const bigPath = join(tmp, "大きい売上.xlsx");
writeFileSync(bigPath, XLSX.write(bigWb, { bookType: "xlsx", type: "buffer" }));

// 表示が出た瞬間を取りこぼさないよう、変化を監視してから読み込ませる
await page.evaluate(() => {
  window.__seen = [];
  const box = document.getElementById("loading");
  const rec = () => { if (!box.hidden) window.__seen.push(document.getElementById("loadStage").textContent); };
  new MutationObserver(rec).observe(box, { attributes: true, subtree: true, childList: true, characterData: true });
  rec();
});
await page.setInputFiles("#fileInput", bigPath);
await page.waitForFunction(() => window.__app.S.fileName === "大きい売上.xlsx", { timeout: 30000 });
const seen = await page.evaluate(() => window.__seen);
check("読み込み中の表示が出る", seen.length > 0, seen.slice(0, 4).join(" → "));
check("「解析中」の段階が出る", seen.some((s) => /解析中/.test(s)), seen.join(" → "));
check("ファイル名が出る", (await page.textContent("#loadName")) === "大きい売上.xlsx");
check("終わったら表示が消える", await page.evaluate(() => document.getElementById("loading").hidden));
check("大きいブックも読める（4万行）",
  (await page.evaluate(() => window.__app.S.sheets[0].usedRows)) === 40001,
  String(await page.evaluate(() => window.__app.S.sheets[0].usedRows)));

// 解析できないファイル（画像など）は、はっきりエラーとして出す
const junkPath = join(tmp, "こわれたブック.xlsx");
writeFileSync(junkPath, Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4, 5, 6, 7, 8]));
await page.setInputFiles("#fileInput", junkPath);
await page.waitForSelector("#loading:not([hidden]) .acts:not([hidden])", { timeout: 10000 });
check("壊れたファイルはエラーとして出る",
  /解析できませんでした/.test(await page.textContent("#loadStage")),
  await page.textContent("#loadStage"));
check("エラーの理由も出す", /PNG|spreadsheet/i.test(await page.textContent("#loadSub")),
  await page.textContent("#loadSub"));
check("エラー時は閉じるボタンが出る", await page.isVisible("#loadClose"));
await page.click("#loadClose");
check("閉じるで消える", await page.evaluate(() => document.getElementById("loading").hidden));

// 中身が空のファイルは、SheetJS が例外を出さず 1 セルのブックとして通してしまう。
// 黙って空の表を出さず、警告として伝えること
const emptyPath = join(tmp, "からっぽ.xlsx");
writeFileSync(emptyPath, "");
await page.setInputFiles("#fileInput", emptyPath);
await page.waitForSelector("#loading:not([hidden]) .acts:not([hidden])", { timeout: 10000 });
check("中身が空のファイルは警告として伝える",
  /データが1件も見つかりませんでした/.test(await page.textContent("#loadStage")),
  await page.textContent("#loadStage"));
check("空ファイルで既存のブックを失わない",
  (await page.evaluate(() => window.__app.S.fileName)) === "大きい売上.xlsx",
  await page.evaluate(() => window.__app.S.fileName));
await page.click("#loadClose");

// 拡張子が対象外なら、読み込む前にはっきり伝える
const txtPath = join(tmp, "メモ.pdf");
writeFileSync(txtPath, "dummy");
await page.setInputFiles("#fileInput", txtPath);
await page.waitForSelector("#loading:not([hidden])", { timeout: 5000 });
check("対象外の拡張子はその場で伝える",
  /Excelファイルではないようです/.test(await page.textContent("#loadStage")),
  await page.textContent("#loadStage"));
check("読み込み済みのブックは壊されない",
  (await page.evaluate(() => window.__app.S.fileName)) === "大きい売上.xlsx");
await page.click("#loadClose");

// あとの検証のためサンプルに戻す
await page.click("#btnSample");
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9d. Excel で開いたときと同じ見え方（色・非表示・幅・固定） ----------
await page.setInputFiles("#fileInput", join(root, "test/fixtures/書式つき.xlsx"));
await page.waitForFunction(() => window.__app.S.fileName === "書式つき.xlsx", { timeout: 20000 });

// 塗りつぶし色がそのまま出る
const fillA1 = await page.$eval('#srcGrid .gc[data-r="0"][data-c="0"]', (n) => getComputedStyle(n).backgroundColor);
check("見出しの塗りつぶし色が出る", fillA1 === "rgb(31, 95, 168)", fillA1);
const fillB3 = await page.$eval('#srcGrid .gc[data-r="2"][data-c="1"]', (n) => getComputedStyle(n).backgroundColor);
check("条件つき色分けのセルも色が出る", fillB3 === "rgb(255, 242, 204)", fillB3);
const textA1 = await page.$eval('#srcGrid .gc[data-r="0"][data-c="0"]', (n) => getComputedStyle(n).color);
check("濃い背景では文字を白くする", textA1 === "rgb(255, 255, 255)", textA1);

// 非表示の列 C と行 4 は出さない
check("非表示の列は描かれない", (await page.locator('#srcGrid .gc[data-c="2"]').count()) === 0);
check("非表示の行は描かれない", (await page.locator('#srcGrid .gc[data-r="3"]').count()) === 0);
check("非表示があることを知らせる", await page.isVisible("#hiddenNote"));
check("シート情報に非表示の数が出る",
  /非表示 1行\/1列/.test(await page.textContent("#srcSheetName")), await page.textContent("#srcSheetName"));
check("非表示の跡に目印が出る", (await page.locator("#srcGrid .after-hidden-c").count()) > 0);

// 「非表示も表示」で見えるようになる
await page.check("#optShowHidden");
check("非表示も表示にすると出てくる",
  (await page.locator('#srcGrid .gc[data-c="2"]').count()) > 0
  && (await page.locator('#srcGrid .gc[data-r="3"]').count()) > 0);
await page.uncheck("#optShowHidden");

// 列幅がファイルどおり（A=14文字ぶん、B=9文字ぶん）
const wA = await page.$eval('#srcGrid .gc[data-r="1"][data-c="0"]', (n) => n.getBoundingClientRect().width);
const wB = await page.$eval('#srcGrid .gc[data-r="1"][data-c="1"]', (n) => n.getBoundingClientRect().width);
check("列ごとに幅が変わる", Math.round(wA) === 84 && Math.round(wB) === 54, `A=${Math.round(wA)} B=${Math.round(wB)}`);

// ウィンドウ枠の固定
await selectRange("B3", "B3");
await page.click("#btnFreeze");
check("固定するとボタンが解除に変わる", (await page.textContent("#btnFreeze")) === "固定を解除");
check("固定した行・列の層が出る",
  (await page.isVisible("#srcGrid .gfrz-top")) && (await page.isVisible("#srcGrid .gfrz-left")));
// スクロールしても固定部分が画面内に残る
await page.evaluate(() => { const w = document.getElementById("srcGrid"); w.scrollTop = 400; w.scrollLeft = 300; });
await page.waitForTimeout(120);
const frozenStuck = await page.evaluate(() => {
  const w = document.getElementById("srcGrid").getBoundingClientRect();
  const t = document.querySelector("#srcGrid .gfrz-top").getBoundingClientRect();
  const l = document.querySelector("#srcGrid .gfrz-left").getBoundingClientRect();
  return { topIn: t.top >= w.top - 1 && t.top < w.top + 40, leftIn: l.left >= w.left - 1 && l.left < w.left + 60 };
});
check("スクロールしても固定した行が上に残る", frozenStuck.topIn, JSON.stringify(frozenStuck));
check("スクロールしても固定した列が左に残る", frozenStuck.leftIn, JSON.stringify(frozenStuck));
// 固定した行・列の「見出し」も一緒に残る
const frozenHeads = await page.evaluate(() => {
  const w = document.getElementById("srcGrid").getBoundingClientRect();
  const ns = [...document.querySelectorAll("#srcGrid .gfrzhead .gh")];
  return ns.map((n) => ({
    label: n.textContent,
    inView: n.getBoundingClientRect().top >= w.top - 1 && n.getBoundingClientRect().left >= w.left - 1,
  }));
});
check("固定した行番号・列名も画面内に残る",
  frozenHeads.length === 3 && frozenHeads.every((h) => h.inView),
  frozenHeads.map((h) => h.label + (h.inView ? "○" : "×")).join(","));
await page.click("#btnFreeze");
check("固定を解除できる", (await page.textContent("#btnFreeze")) === "ここで固定");

// ---- 9d-2. 見出しを押すと固定メニューが出る ------------------------------
await page.evaluate(() => { const s = window.__app.S.sheets[window.__app.S.active]; s.freeze = null; });
await page.evaluate(() => { const w = document.getElementById("srcGrid"); w.scrollTop = 0; w.scrollLeft = 0; });
await page.waitForTimeout(80);
const rowHead3 = await page.locator('#srcGrid .growh .gh[data-r="2"]').boundingBox();
await page.mouse.click(rowHead3.x + 20, rowHead3.y + 10);
check("行見出しを押すとメニューが出る", await page.isVisible(".hmenu"));
check("メニューに対象の行が出る", (await page.textContent(".hmenu .ttl")) === "3行目",
  await page.textContent(".hmenu .ttl"));
check("行を押したときは上で固定と出る",
  /この行の上で固定/.test(await page.textContent(".hmenu")), await page.textContent(".hmenu"));
await page.click('.hmenu button:has-text("この行の上で固定")');
check("メニューから固定できる",
  (await page.evaluate(() => window.__app.S.sheets[window.__app.S.active].freeze)).r === 2,
  JSON.stringify(await page.evaluate(() => window.__app.S.sheets[window.__app.S.active].freeze)));
check("固定するとメニューが閉じる", (await page.locator(".hmenu").count()) === 0);

// 列見出しからも同じように固定できる
const colHeadC = await page.locator('#srcGrid .gcolh .gh[data-c="1"]').boundingBox();
await page.mouse.click(colHeadC.x + 20, colHeadC.y + 10);
check("列見出しでは左で固定と出る",
  /この列の左で固定/.test(await page.textContent(".hmenu")), await page.textContent(".hmenu"));
check("固定中は解除も選べる", /固定を解除/.test(await page.textContent(".hmenu")));
await page.click('.hmenu button:has-text("この列の左で固定")');
const fz2 = await page.evaluate(() => window.__app.S.sheets[window.__app.S.active].freeze);
check("行と列の固定は両立する", fz2.r === 2 && fz2.c === 1, JSON.stringify(fz2));

// 固定した行・列は画面に貼り付いているので、スクロール後もその見出しを正しく掴めること
await page.evaluate(() => { const w = document.getElementById("srcGrid"); w.scrollTop = 300; w.scrollLeft = 200; });
await page.waitForTimeout(120);
const frozenHead0 = await page.locator('#srcGrid .gfrzhead .gh[data-r="0"]').boundingBox();
await page.mouse.click(frozenHead0.x + 20, frozenHead0.y + 10);
check("スクロール後も固定行の見出しを正しく掴める",
  (await page.textContent(".hmenu .ttl")) === "1行目", await page.textContent(".hmenu .ttl"));
check("1行目では固定できないと伝える",
  /1行目より上は固定できません/.test(await page.textContent(".hmenu")), await page.textContent(".hmenu"));
await page.keyboard.press("Escape");
check("Escでメニューが閉じる", (await page.locator(".hmenu").count()) === 0);
// 固定セルそのものも正しく選べる
const frozenCell = await page.locator('#srcGrid .gfrz-corner .gc[data-r="1"][data-c="0"]').boundingBox();
await page.mouse.click(frozenCell.x + 10, frozenCell.y + 10);
check("スクロール後も固定セルを正しく選べる",
  /A2/.test(await page.textContent("#selInfo")), await page.textContent("#selInfo"));
await page.evaluate(() => { const w = document.getElementById("srcGrid"); w.scrollTop = 0; w.scrollLeft = 0; });
await page.evaluate(() => { window.__app.S.sheets[window.__app.S.active].freeze = null; window.__app.S.sel = null; });

// ---- 9d-3. 元データの全画面表示 -----------------------------------------
const beforeW = await page.evaluate(() => document.getElementById("srcGrid").clientWidth);
await page.click("#btnMaxSrc");
await page.waitForTimeout(200);
const afterW = await page.evaluate(() => document.getElementById("srcGrid").clientWidth);
check("全画面で元データが広がる", afterW > beforeW * 1.8, `${beforeW}px → ${afterW}px`);
check("全画面では他のペインが隠れる",
  !(await page.isVisible("#dstGrid")) && !(await page.isVisible("#sheetList")));
check("ボタンが戻す表示になる", (await page.textContent("#btnMaxSrc")) === "⤡ 戻す");
check("全画面でもセルは描かれる",
  (await page.$$eval("#srcGrid .gc", (ns) => ns.filter((n) => n.textContent.trim()).length)) > 5);
await page.keyboard.press("Escape");
await page.waitForTimeout(200);
check("Escで元の3画面に戻る",
  (await page.isVisible("#dstGrid")) && (await page.textContent("#btnMaxSrc")) === "⤢ 全画面");

// ---- 9e. 末尾へ移動（置かずに表示だけ動かす） ---------------------------
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await selectRange("A2", "F2");
await page.click("#btnAppendDown");
check("移動ボタンでは置かれない（ブロックは増えない）",
  (await page.locator("#blockList .block-card").count()) === 0,
  String(await page.locator("#blockList .block-card").count()));
check("移動先に目印が残る", (await page.locator("#dstGrid .dropghost").count()) === 1);
check("空のときの移動先は A1",
  /^A1 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));

// 目印のところへドラッグして置く（選択範囲は掴んだままなので、そのまま運べる）
let ghostBox = await page.locator("#dstGrid .dropghost").boundingBox();
let selBox = await page.locator("#srcGrid .selbox").boundingBox();
await page.mouse.move(selBox.x + selBox.width / 2, selBox.y + selBox.height / 2);
await page.mouse.down();
await page.mouse.move(ghostBox.x + 20, ghostBox.y + 10, { steps: 12 });
await page.mouse.move(ghostBox.x + 22, ghostBox.y + 12, { steps: 3 });
await page.mouse.up();
check("選択範囲の中を掴んでドラッグできる",
  (await page.locator("#blockList .block-card").count()) === 1,
  String(await page.locator("#blockList .block-card").count()));
check("目印の位置に置かれる",
  (await page.textContent("#blockList .block-card .row2 .to")) === "抜粋1!A1",
  await page.textContent("#blockList .block-card .row2 .to"));
check("置いたら目印は消える", (await page.locator("#dstGrid .dropghost").count()) === 0);

// 末尾の右へ移動 → その位置へドラッグ
await selectRange("A3", "F3");
await page.click("#btnAppendRight");
check("末尾の右は直前のブロックの右隣",
  /^G1 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));
ghostBox = await page.locator("#dstGrid .dropghost").boundingBox();
selBox = await page.locator("#srcGrid .selbox").boundingBox();
await page.mouse.move(selBox.x + selBox.width / 2, selBox.y + selBox.height / 2);
await page.mouse.down();
await page.mouse.move(ghostBox.x + 20, ghostBox.y + 10, { steps: 12 });
await page.mouse.move(ghostBox.x + 22, ghostBox.y + 12, { steps: 3 });
await page.mouse.up();
const appended = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("末尾の下・右に並べられる", appended.join(",") === "抜粋1!A1,抜粋1!G1", appended.join(","));

// 置き場所を間違えたら、ブロックのどこを掴んでも動かせる
// （直前の移動で出力側が動いているので、先頭に戻してから掴む）
await page.evaluate(() => {
  const w = document.getElementById("dstGrid");
  w.scrollTop = 0; w.scrollLeft = 0;
});
await page.waitForTimeout(120);
const blockBox = await page.locator('#dstGrid .blockbox').first().boundingBox();
const moveTo = await page.locator('#dstGrid .gc[data-r="5"][data-c="1"]').boundingBox();
await page.mouse.move(blockBox.x + blockBox.width / 2, blockBox.y + blockBox.height / 2);
await page.mouse.down();
await page.mouse.move(moveTo.x + 20, moveTo.y + 10, { steps: 14 });
await page.mouse.move(moveTo.x + 22, moveTo.y + 12, { steps: 3 });
await page.mouse.up();
const movedTo = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("ブロックの本体を掴んで置き直せる", movedTo.includes("抜粋1!B6"), movedTo.join(","));
check("ブロックは増えない（移動であって複製ではない）",
  (await page.locator("#blockList .block-card").count()) === 2,
  String(await page.locator("#blockList .block-card").count()));

// 選択の中をクリックだけしたら、そのセルへ畳む（Excel と同じ）
const selBox2 = await page.locator("#srcGrid .selbox").boundingBox();
await page.mouse.click(selBox2.x + 10, selBox2.y + 10);
const collapsed = await page.textContent("#selInfo");
check("選択の中をクリックすると1セルになる", /A3:A3|A3\b/.test(collapsed) && /1×1/.test(collapsed), collapsed);

// 1行目に置いたときも、ラベルが列見出しに隠れず全部読めること
await page.evaluate(() => window.__app.runScript("シート追加 先頭\n書式つきのA2:F2を先頭のA1に置く", true));
await page.click('#dstGrid .blockbox');
await page.waitForTimeout(100);
const topTag = await page.evaluate(() => {
  const box = document.querySelector("#dstGrid .blockbox");
  const tag = box.querySelector(".tag");
  const grid = document.getElementById("dstGrid").getBoundingClientRect();
  const head = document.querySelector("#dstGrid .gcolh").getBoundingClientRect();
  const t = tag.getBoundingClientRect();
  return {
    below: tag.classList.contains("below"),
    fullyVisible: t.top >= grid.top - 0.5 && t.bottom <= grid.bottom + 0.5,
    belowHeader: t.top >= head.bottom - 0.5,
    height: Math.round(t.height),
  };
});
check("1行目のラベルは下側に出る", topTag.below, JSON.stringify(topTag));
check("1行目のラベルが見出しに隠れない", topTag.belowHeader && topTag.fullyVisible, JSON.stringify(topTag));

// 上に余裕がある位置なら、これまでどおり上に出る
await page.evaluate(() => {
  window.__app.runScript("シート追加 途中\n書式つきのA2:F2を途中のA6に置く", true);
});
await page.click('#dstGrid .blockbox');
await page.waitForTimeout(100);
const midTag = await page.evaluate(() => {
  const tag = document.querySelector("#dstGrid .blockbox .tag");
  const grid = document.getElementById("dstGrid").getBoundingClientRect();
  const t = tag.getBoundingClientRect();
  return { below: tag.classList.contains("below"), fullyVisible: t.top >= grid.top - 0.5 };
});
check("余裕があればラベルは上のまま", !midTag.below && midTag.fullyVisible, JSON.stringify(midTag));

// 元データ側も同じ（1行目を選んだときのつまみ）
await selectRange("A1", "C1");
const grabChip = await page.evaluate(() => {
  const g = document.querySelector("#srcGrid .selbox .grab");
  const grid = document.getElementById("srcGrid").getBoundingClientRect();
  const head = document.querySelector("#srcGrid .gcolh").getBoundingClientRect();
  const t = g.getBoundingClientRect();
  return { below: g.classList.contains("below"), belowHeader: t.top >= head.bottom - 0.5 };
});
check("元データの1行目のつまみも下側に出る", grabChip.below && grabChip.belowHeader, JSON.stringify(grabChip));
// 移動先が画面の外なら、そこまで出力側の表示が動く
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript("シート追加 遠い\n書式つきのA2:F2を遠いのA60に置く", true);
  document.getElementById("dstGrid").scrollTop = 0;
});
await selectRange("A5", "F5");
await page.click("#btnAppendDown");
await page.waitForTimeout(150);
const moved = await page.evaluate(() => {
  const w = document.getElementById("dstGrid");
  const g = w.querySelector(".dropghost").getBoundingClientRect();
  const view = w.getBoundingClientRect();
  return { scrollTop: w.scrollTop, inView: g.top >= view.top - 1 && g.bottom <= view.bottom + 1 };
});
check("移動先が画面外ならそこまで表示が動く", moved.scrollTop > 200, `scrollTop=${Math.round(moved.scrollTop)}`);
check("移動先の目印が画面内に入る", moved.inView, JSON.stringify(moved));
// 端ぎりぎりではなく、真ん中あたりに来ること
const centered = await page.evaluate(() => {
  const w = document.getElementById("dstGrid").getBoundingClientRect();
  const g = document.querySelector("#dstGrid .dropghost").getBoundingClientRect();
  return { offset: Math.abs((g.top + g.height / 2) - (w.top + w.height / 2)), viewH: w.height };
});
check("移動先は画面の真ん中あたりに来る", centered.offset < centered.viewH * 0.25,
  `中心から ${Math.round(centered.offset)}px（画面高 ${Math.round(centered.viewH)}px）`);
check("移動先は末尾の下 A61",
  /^A61 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));

// 文章で指示したときは、置いたブロックまで表示が動いて光る
await page.click("#tabCmd");     // 手順書タブから1行指示に戻す
await page.fill("#cmdInput", "書式つきのA2:F2を遠いのA80に置く");
await page.press("#cmdInput", "Enter");
await page.waitForTimeout(200);
check("文章で置いたときも表示が追いかける",
  (await page.evaluate(() => document.getElementById("dstGrid").scrollTop)) > 1200,
  String(Math.round(await page.evaluate(() => document.getElementById("dstGrid").scrollTop))));
check("置いた直後は光って知らせる",
  (await page.locator("#dstGrid .blockbox.flash").count()) === 1);

// あとの検証のためサンプルに戻す
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await page.click("#btnSample");
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9f. 末尾は「値の入っている範囲」で判断する -------------------------
// 行まるごとの選択は使用範囲いっぱい（末尾は空欄だらけ）になる。
// 空欄まで末尾に数えると、ずっと右の何も無い場所へ飛んでしまう
await page.setInputFiles("#fileInput", join(root, "test/fixtures/横長.xlsx"));
await page.waitForFunction(() => window.__app.S.fileName === "横長.xlsx", { timeout: 20000 });
const sparse = await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.setSelection(4, 0, 4, 79);                 // 5行目まるごと（値は A〜C の3列だけ）
  const o = A.S.out[0] || A.S.out[A.S.activeOut];
  A.addBlock({ sheet: 0, r1: 4, c1: 0, r2: 4, c2: 79 }, 0, 0, o);
  return A.S.sel.c2;
});
check("行まるごとの選択は使用範囲いっぱいになる", sparse === 79, String(sparse));
await page.click("#btnAppendRight");
check("末尾の右は値の右隣（空欄は数えない）",
  /^D1 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));
check("遠くまで飛ばない",
  (await page.evaluate(() => document.getElementById("dstGrid").scrollLeft)) < 300,
  String(Math.round(await page.evaluate(() => document.getElementById("dstGrid").scrollLeft))));
await page.click("#btnAppendDown");
check("末尾の下も値の下（空欄は数えない）",
  /^A2 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));

// 本当に全列に値がある行なら、これまでどおり右端の続きへ
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.setSelection(0, 0, 0, 79);
  A.addBlock({ sheet: 0, r1: 0, c1: 0, r2: 0, c2: 79 }, 0, 0, A.S.out[0] || A.S.out[A.S.activeOut]);
});
await page.click("#btnAppendRight");
check("値が全列にある行では右端の続きへ行く",
  /^CC1 へ/.test(await page.textContent("#dstGrid .dropghost .lbl")),
  await page.textContent("#dstGrid .dropghost .lbl"));

await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await page.click("#btnSample");
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

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
