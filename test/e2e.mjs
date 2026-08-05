/**
 * dist/index.html をブラウザで動かして、選択 → 配置 → 生成まで通しで検証する。
 *   node test/e2e.mjs
 * Playwright と SheetJS はグローバル/ローカルどちらでも解決する。
 */
import { readFileSync, writeFileSync, mkdirSync, rmSync } from "node:fs";
import { execFileSync } from "node:child_process";
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

/** サンプルだけを開いた状態に戻す。最初の画面を閉じたあとはボタンが無いので API で読む */
async function loadSampleAgain() {
  await page.evaluate(() => {
    const A = window.__app;
    while (A.S.files.length) A.removeFile(0);
    A.loadSample();
  });
  await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");
}

/** そのファイルだけを開いた状態にする（複数開けるようになったので、前のは閉じる） */
async function openOnly(path, name, timeout) {
  await page.evaluate(() => { const A = window.__app; while (A.S.files.length) A.removeFile(0); });
  await page.setInputFiles("#fileInput", path);
  await page.waitForFunction((n) => window.__app.S.fileName === n, name, { timeout: timeout || 20000 });
}

/** いま選んでいる範囲を "売上明細!A2:D6 5×4" の形で返す（画面には出さなくなったため） */
async function selInfo() {
  return await page.evaluate(() => {
    const A = window.__app, g = A.S.sel;
    if (!g) return "未選択";
    const col = (c) => { let s = "", n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
    return A.S.sheets[g.sheet].name + "!" + col(g.c1) + (g.r1 + 1) + ":" + col(g.c2) + (g.r2 + 1)
      + "  " + (g.r2 - g.r1 + 1) + "×" + (g.c2 - g.c1 + 1);
  });
}

/** 範囲を選ぶ（画面の入力欄は廃止したので、アプリの API で指定する） */
async function selectRange(from, to) {
  await page.evaluate(([f, t]) => {
    const A = window.__app, a = A.parseCell(f), b = A.parseCell(t);
    A.setSelection(a.r, a.c, b.r, b.c);
    A.revealSelection();
  }, [from, to]);
}

/**
 * 出力シートのブロックを、その中心を直接押して選ぶ。
 * page.click は「見えているか」の判定でスクロールしてしまうことがあり、
 * ラベルの上下（見出しに隠れないための切り替え）を測る妨げになる。
 */
async function clickBlock() {
  const b = await page.locator("#dstGrid .blockbox").first().boundingBox();
  await page.mouse.click(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(100);
}

/** いま選んでいる範囲を、出力シートの指定セル（例 "G1"）へドラッグして置く */
async function dragSelectionToRef(ref) {
  const at = await page.evaluate((r) => {
    const p = window.__app.parseCell(r);
    return { r: p.r, c: p.c };
  }, ref);
  const target = await page.locator(`#dstGrid .gc[data-r="${at.r}"][data-c="${at.c}"]`).boundingBox();
  const sel = await page.locator("#srcGrid .selbox").boundingBox();
  await page.mouse.move(sel.x + sel.width / 2, sel.y + sel.height / 2);
  await page.mouse.down();
  await page.mouse.move(target.x + 20, target.y + 10, { steps: 12 });
  await page.mouse.move(target.x + 22, target.y + 12, { steps: 3 });
  await page.mouse.up();
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
check("右上にサンプルのボタンは無い", (await page.locator("#btnSample").count()) === 0);
check("最初の画面にはサンプルのボタンがある", await page.isVisible("#btnSample2"));
check("サンプルのダウンロードは出さない", (await page.locator("#btnSampleDl").count()) === 0);
check("1・2・3 の手順書きは出さない", (await page.locator(".dropzone .steps").count()) === 0);
await page.click("#btnSample2");
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
let selText = await selInfo();
check("ドラッグで A2:D6 が選択される", /A2:D6/.test(selText) && /5×4/.test(selText), selText);

// ---- 3. 選択範囲を出力シートへドラッグ＆ドロップ ------------------------
const grab = await page.locator("#srcGrid .selbox").boundingBox();
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

// ---- 4. 範囲を指定しての選択 + ドラッグでの配置 --------------------------
await selectRange("A1", "C3");
selText = await selInfo();
check("A1:C3 を選択できる", /A1:C3/.test(selText), selText);
check("「全体を選択」ボタンは無い", (await page.locator("#btnSelAll").count()) === 0);
check("「出力へ配置」ボタンも無い", (await page.locator("#btnPlace").count()) === 0);
await dragSelectionToRef("A13");
blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("ドラッグでの配置でブロックが2件になる", blocks.length === 2, blocks.join(" / "));
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("2件目の配置先は A13", dest[1] === "抜粋1!A13", dest.join(","));

// ---- 4b. 出力へ持っていった範囲は元データ側に残る ------------------------
// 別のシートを見て戻ってきても消えないこと、色が出力側のブロックと同じことを確かめる
const usedFrames = async () => await page.evaluate(() =>
  [...document.querySelectorAll("#srcGrid .usedbox")].map((n) => ({
    id: n.dataset.id, bc: n.style.getPropertyValue("--bc"),
    lbl: n.querySelector(".tag").textContent })));
const outColors = async () => await page.evaluate(() =>
  [...document.querySelectorAll("#dstGrid .blockbox")].map((n) => ({
    id: n.dataset.id, bc: n.style.getPropertyValue("--bc") })));
const used1 = await usedFrames(), outs1 = await outColors();
check("持っていった範囲の枠が元データに残る", used1.length === 2, JSON.stringify(used1));
check("枠の色は出力側のブロックと同じ",
  used1.every((u) => (outs1.find((o) => o.id === u.id) || {}).bc === u.bc),
  JSON.stringify({ used: used1, out: outs1 }));
check("行き先がラベルに入っている",
  used1.some((u) => /→ 抜粋1!A13/.test(u.lbl)), JSON.stringify(used1.map((u) => u.lbl)));
// 別のシートへ行くと、そのシートのぶんだけになる
await page.click('#srcTabs .tab:nth-child(2)');
await page.waitForTimeout(200);
check("別のシートには別のシートぶんだけ出る", (await usedFrames()).length === 0,
  JSON.stringify(await usedFrames()));
// 戻ってくると、また出る（前は消えたままだった）
await page.click('#srcTabs .tab:nth-child(1)');
await page.waitForTimeout(200);
const used2 = await usedFrames();
check("戻ってくると枠も戻る", used2.length === 2, JSON.stringify(used2));
check("戻っても色は変わらない",
  JSON.stringify(used2.map((u) => u.bc)) === JSON.stringify(used1.map((u) => u.bc)),
  JSON.stringify(used2.map((u) => u.bc)));
await page.screenshot({
  path: join(root, "test/shots/pairs.png"),
  clip: { x: 232, y: 48, width: 1208, height: 480 },
});

// ---- 5. 文章コマンド ----------------------------------------------------
await page.evaluate(() => window.__app.runScript("支店別サマリの2行目から7行目を抜粋1のF3に置く", false));
blocks = await page.$$eval("#blockList .block-card .src", (ns) => ns.map((n) => n.textContent));
check("文章コマンドで3件目が追加される", blocks.length === 3, blocks.join(" / "));
check("コマンドの範囲解釈が正しい", blocks[2] === "支店別サマリ!A2:C7", blocks[2]);
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("コマンドの配置先が F3", dest[2] === "抜粋1!F3", dest[2]);

// シート追加コマンド
await page.evaluate(() => window.__app.runScript("シート追加 集計用", false));
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
// 出力側の下の帯は、選んだ範囲の集計だけにした
check("件数や貼り付け位置の欄は出さない",
  (await page.locator("#outStat").count()) === 0
  && (await page.locator("#inspAnchor").count()) === 0
  && (await page.locator("#inspTitle").count()) === 0
  && (await page.locator("#inspEmpty").count()) === 0);
// 向きと削除はツールバーへ。選ぶまでは押せない
check("ブロックを選ぶ前は向きを変えられない", await page.isDisabled("#inspTrans"));
check("向きの切り替えは保持の設定と同じツールバーにいる", await page.evaluate(() =>
  document.querySelector(".pane-dst .toolbar").contains(document.getElementById("inspTrans"))));
check("削除ボタンは無い（Delete キーで外す）", (await page.locator("#inspDel").count()) === 0);
await page.click("#blockList .block-card >> nth=0");
check("ブロックを選ぶと押せるようになる", !(await page.isDisabled("#inspTrans")));
await page.click("#inspTrans");
blocks = await page.$$eval("#blockList .block-card .row2", (ns) => ns.map((n) => n.textContent));
check("転置が効く", /転置/.test(blocks[0]), blocks[0]);
await page.keyboard.press("Control+z");
dest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("Ctrl+Z で戻せる", dest[0] === "抜粋1!B3", dest[0]);

const before = await page.$$eval("#blockList .block-card", (ns) => ns.length);
await page.click("#blockList .block-card >> nth=0 >> .icon-btn");
const after = await page.$$eval("#blockList .block-card", (ns) => ns.length);
check("ブロックを削除できる", after === before - 1, `${before} → ${after}`);

// ---- 7b. 範囲を選んで Delete でまとめて外す ------------------------------
// 3 つ置いて、そのうち 2 つに掛かる範囲を選んで Delete
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null; A.clearDstSel();
  A.runScript(`シート追加 まとめて
売上明細のA1:C3をまとめてのA1に置く
売上明細のA1:C3をまとめてのA5に置く
売上明細のA1:C3をまとめてのA20に置く`, true);
});
await page.waitForTimeout(250);
check("3つ置いた", (await page.$$eval("#blockList .block-card", (ns) => ns.length)) === 3);
// 出力グリッドを触ってから、A1:C8（上の 2 つに掛かる）を選ぶ
await page.click('#dstGrid .gc[data-r="10"][data-c="0"]');
await page.evaluate(() => window.__app.setDstSel(0, 0, 7, 2));
await page.waitForTimeout(120);
await page.keyboard.press("Delete");
await page.waitForTimeout(200);
const leftAfterDel = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("選んだ範囲に掛かるブロックだけ外れる",
  leftAfterDel.join(",") === "まとめて!A20", leftAfterDel.join(","));
check("何件外したか知らせる",
  /ブロック 2 件を外しました/.test(await page.textContent("#toasts")),
  (await page.textContent("#toasts")).slice(0, 60));
await page.keyboard.press("Control+z");
await page.waitForTimeout(200);
check("Ctrl+Z でまとめて戻せる",
  (await page.$$eval("#blockList .block-card", (ns) => ns.length)) === 3,
  String(await page.$$eval("#blockList .block-card", (ns) => ns.length)));

// 何も無い範囲では消さずに知らせる
await page.evaluate(() => { document.getElementById("toasts").innerHTML = ""; });
await page.click('#dstGrid .gc[data-r="10"][data-c="0"]');
await page.evaluate(() => window.__app.setDstSel(9, 0, 12, 2));
await page.keyboard.press("Delete");
await page.waitForTimeout(200);
check("何も無い範囲では消えない",
  (await page.$$eval("#blockList .block-card", (ns) => ns.length)) === 3);
check("何も無いことを知らせる",
  /選んだ範囲にブロックはありません/.test(await page.textContent("#toasts")),
  (await page.textContent("#toasts")).slice(0, 60));

// 元データ側にいるときは効かない（うっかり消さない）
await page.evaluate(() => { window.__app.setDstSel(0, 0, 7, 2); });
await page.click('#srcGrid .gc[data-r="3"][data-c="0"]');
await page.keyboard.press("Delete");
await page.waitForTimeout(200);
check("元データ側での Delete では消えない",
  (await page.$$eval("#blockList .block-card", (ns) => ns.length)) === 3);

// 範囲を選んでいなくても、選択中のブロックは Delete で外せる
await page.click('#dstGrid .blockbox >> nth=0');
await page.evaluate(() => window.__app.clearDstSel());
await page.waitForTimeout(120);
await page.keyboard.press("Delete");
await page.waitForTimeout(200);
check("選んでいるブロックは Delete で外せる",
  (await page.$$eval("#blockList .block-card", (ns) => ns.length)) === 2,
  String(await page.$$eval("#blockList .block-card", (ns) => ns.length)));
await page.keyboard.press("Control+z");
await page.waitForTimeout(150);

// 元の状態に戻す
await page.evaluate(() => { window.__app.clearDstSel(); });
await loadSampleAgain();
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(`シート追加 抜粋1
売上明細のA2:D6を抜粋1のB3に置く
売上明細のA1:C3を抜粋1のA13に置く
支店別サマリのA2:C7を抜粋1のF3に置く
シート追加 集計用
商品マスタのA1:C7を集計用のB2に置く`, true);
});
await page.waitForTimeout(250);
await page.click("#dstTabsHost .tab >> nth=0");

// ---- 7c. Python への書き出し（同じ処理を Snowflake などで動かす） --------
// 画面が書き出す xlsx と、書き出した Python が作る xlsx が一致することまで見る
await loadSampleAgain();
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(`シート追加 まとめ
売上明細のA1:E8をまとめのA1に置く
支店別サマリのA1:C7をまとめのG1に置く
商品マスタのA1:C4をまとめのA11に置く（転置）`, true);
});
await page.waitForTimeout(250);

await page.evaluate(() => window.__app.openScript(true));
await page.click("#scPy");
await page.waitForTimeout(150);
check("Python の書き出し欄が出る", await page.isVisible("#pyForm"));
check("元ファイルごとに置き場所を書ける",
  (await page.$$eval("#pyInputs .fnm", (ns) => ns.map((n) => n.textContent))).join(",") === "サンプル売上.xlsx",
  (await page.$$eval("#pyInputs .fnm", (ns) => ns.map((n) => n.textContent))).join(","));
// どちらが入力でどちらが出力か、見出しで分かること
check("読み込む側と書き出す側に見出しがつく", await page.evaluate(() =>
  [...document.querySelectorAll("#pyForm .pylabel")].map((n) => n.textContent.trim()).join(" / ")) === "読み込む Excel / 書き出す Excel",
  await page.evaluate(() =>
    [...document.querySelectorAll("#pyForm .pylabel")].map((n) => n.textContent.trim()).join(" / ")));
check("読み込む欄は元ファイルの下にある", await page.evaluate(() => {
  const lab = document.getElementById("pyInLabel");
  const inp = document.querySelector("#pyInputs input");
  return lab.getBoundingClientRect().y <= inp.getBoundingClientRect().bottom;
}));
check("出力ファイル名の既定が入る",
  (await page.inputValue("#pyOut")) === "サンプル売上_抜粋.xlsx", await page.inputValue("#pyOut"));

const pyDir = join(tmp, "py");
mkdirSync(pyDir, { recursive: true });
await page.fill("#pyOut", "by-python.xlsx");
const [dlPy] = await Promise.all([page.waitForEvent("download"), page.click("#pyOk")]);
check("Python として書き出せる", dlPy.suggestedFilename() === "by-python.py", dlPy.suggestedFilename());
await dlPy.saveAs(join(pyDir, "extract.py"));
const pySrc = readFileSync(join(pyDir, "extract.py"), "utf8");
check("入出力が先頭にまとまっている",
  /SOURCES = \{/.test(pySrc) && /OUTPUT = "by-python\.xlsx"/.test(pySrc), pySrc.slice(0, 60));
check("配置が手順として並ぶ",
  (pySrc.match(/^\s+\{"out":/gm) || []).length === 3,
  String((pySrc.match(/^\s+\{"out":/gm) || []).length));
check("転置も引き継がれる", /"transpose": True/.test(pySrc));
check("openpyxl 以外は要らない",
  !/^import (?!openpyxl|from)/m.test(pySrc) && /import openpyxl/.test(pySrc));
check("後処理へつなぐ入口がある", /def to_dataframes\(/.test(pySrc));

// 実際に走らせて、画面の書き出しと突き合わせる（python3 と openpyxl があるときだけ）
let pyReady = true;
try { execFileSync("python3", ["-c", "import openpyxl"], { stdio: "ignore" }); } catch { pyReady = false; }
if (!pyReady) {
  console.log("  --   python3 / openpyxl が無いので、走らせての照合は省略");
} else {
  // 元データと、画面が書き出した xlsx を並べる
  const uiPath = join(pyDir, "サンプル売上.xlsx");
  writeFileSync(uiPath, Buffer.from(XLSX.write(await page.evaluate(() => {
    const wb = window.__app.S.files[0].wb;
    return { SheetNames: wb.SheetNames, Sheets: wb.Sheets };
  }), { bookType: "xlsx", type: "buffer" })));
  await page.evaluate(() => window.__app.openScript(false));
  const [dlUi] = await Promise.all([page.waitForEvent("download"), page.click("#btnGen")]);
  await dlUi.saveAs(join(pyDir, "by-ui.xlsx"));

  let ran = "";
  try { ran = execFileSync("python3", ["extract.py"], { cwd: pyDir, encoding: "utf8" }); }
  catch (e) { ran = "ERROR " + (e.stderr || e.message); }
  check("書き出した Python がそのまま走る", /書き出しました/.test(ran), ran.trim().slice(0, 120));

  const byUi = XLSX.read(readFileSync(join(pyDir, "by-ui.xlsx")), { type: "buffer" });
  const byPy = XLSX.read(readFileSync(join(pyDir, "by-python.xlsx")), { type: "buffer" });
  check("出力シートの顔ぶれが同じ",
    byUi.SheetNames.join(",") === byPy.SheetNames.join(","),
    byUi.SheetNames.join(",") + " / " + byPy.SheetNames.join(","));
  let pyDiff = 0, pyCells = 0, firstDiff = "";
  const normv = (v) => (v instanceof Date ? v.toISOString().slice(0, 10)
    : typeof v === "number" ? Math.round(v * 1e6) / 1e6 : v === undefined ? null : v);
  for (const nm of byUi.SheetNames) {
    const wa = byUi.Sheets[nm], wb2 = byPy.Sheets[nm] || {};
    const rg = XLSX.utils.decode_range(wa["!ref"]);
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      for (let c = rg.s.c; c <= rg.e.c; c++) {
        const ad = XLSX.utils.encode_cell({ r, c });
        pyCells++;
        const va = normv(wa[ad] && wa[ad].v), vb = normv(wb2[ad] && wb2[ad].v);
        if (JSON.stringify(va) !== JSON.stringify(vb)) {
          if (!firstDiff) firstDiff = `${nm}!${ad} ${JSON.stringify(va)} vs ${JSON.stringify(vb)}`;
          pyDiff++;
        }
      }
    }
    const ma = (wa["!merges"] || []).map((m) => XLSX.utils.encode_range(m)).sort().join(",");
    const mb = (wb2["!merges"] || []).map((m) => XLSX.utils.encode_range(m)).sort().join(",");
    check(`結合セルも同じ（${nm}）`, ma === mb, `${ma || "なし"} / ${mb || "なし"}`);
  }
  check("画面の書き出しと 1 セルも食い違わない", pyDiff === 0,
    `${pyCells}セル中 ${pyDiff}件ちがう ${firstDiff}`);
}
await page.evaluate(() => window.__app.openScript(false));
await loadSampleAgain();
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(`シート追加 抜粋1
売上明細のA2:D6を抜粋1のB3に置く
売上明細のA1:C3を抜粋1のA13に置く
支店別サマリのA2:C7を抜粋1のF3に置く
シート追加 集計用
商品マスタのA1:C7を集計用のB2に置く`, true);
});
await page.waitForTimeout(250);
await page.click("#dstTabsHost .tab >> nth=0");

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
await page.click("#btnScript");
check("右上のボタンで手順書が開く", await page.isVisible("#scText"));
check("1行指示の入力欄は無くなった", (await page.locator("#cmdInput").count()) === 0);

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
await page.click("#scClose");                       // 手順書を閉じてから書き出す
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
await page.click("#btnScript");                     // 手順書を開き直す
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
await page.click("#scClose");                       // 本体をさわるので手順書を閉じる
await page.mouse.move(2, 2);                        // どのブロックにも触れていない位置へ
await page.waitForTimeout(80);
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
await page.click("#btnScript");                     // 手順書を開き直す
await page.click("#scFromBlocks");
await page.click("#scRepeat");
// どの行を繰り返すか尋ねる欄が出る
check("繰り返しにすると入力欄が出る", await page.isVisible("#repeatForm"));
check("見つけた行番号が入っている", (await page.inputValue("#repeatValues")) === "10",
  await page.inputValue("#repeatValues"));
check("何をするのか文章で示す",
  /10行目のところを、どの行で繰り返しますか/.test(await page.textContent("#repeatLabel")),
  await page.textContent("#repeatLabel"));
check("行・列・シートから選べる", await page.evaluate(() =>
  [...document.querySelectorAll('input[name="repeatKind"]')].map((n) => n.value).join(",")) === "行,列,シート");
// 繰り返す行を入れて決定
await page.fill("#repeatValues", "10, 12, 17, 20");
await page.click("#repeatOk");
check("決定すると入力欄は閉じる", !(await page.isVisible("#repeatForm")));
const wrapped = await page.inputValue("#scText");
check("繰り返しにするで包まれる",
  /繰り返し 行 = 10, 12, 17, 20\n {2}売上明細のA\{行\}:E\{行\}をまとめのA1に置く\n {2}支店別サマリのA\{行\}:C\{行\}をまとめのF1に置く\nここまで/.test(wrapped),
  wrapped.split("\n").filter((l) => /繰り返し|\{行\}|ここまで/.test(l)).join(" / "));
check("シート追加は繰り返しの外に残る",
  /シート追加 まとめ\n繰り返し/.test(wrapped),
  wrapped.split("\n").filter((l) => /シート追加/.test(l)).join(","));
check("入れた行が繰り返しの値になる", /繰り返し 行 = 10, 12, 17, 20/.test(wrapped),
  wrapped.split("\n").filter((l) => /繰り返し/.test(l)).join(","));
await page.click("#scRun");
await page.waitForTimeout(200);
const wrapDest = await page.$$eval("#blockList .block-card .row2 .to", (ns) => ns.map((n) => n.textContent));
check("包んだ手順書がそのまま4周ぶん動く",
  wrapDest.join(",") === "まとめ!A1,まとめ!F1,まとめ!A2,まとめ!F2,まとめ!A3,まとめ!F3,まとめ!A4,まとめ!F4",
  wrapDest.join(","));

/** 手順書を実行して、置かれた先を「出力!位置」で返す */
const placedByEarly = async (script) => await page.evaluate((sc) => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(sc, true);
  const col = (c) => { let s = "", n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  return A.S.out.flatMap((o) => o.blocks.map((b) => o.name + "!" + col(b.dc) + (b.dr + 1)));
}, script);

// ---- 5b-1. 書き方の一覧 ---------------------------------------------------
await page.evaluate(() => window.__app.openScript(true));
check("一覧は最初は閉じている", !(await page.isVisible("#scCheat")));
await page.click("#scHelp");
await page.waitForTimeout(120);
check("「書き方の一覧」で開く", await page.isVisible("#scCheat"));
check("ボタンの文言が閉じる側になる", (await page.textContent("#scHelp")) === "一覧を閉じる");
const cheatHeads = await page.$$eval(".cheat h4", (ns) => ns.map((n) => n.textContent));
check("種類ごとに分かれている",
  cheatHeads.join(",") === "置く,範囲の書き方,名前をつけて、その隣に置く,繰り返し,シートとファイル,その他",
  cheatHeads.join(","));
check("ひととおりの書き方が並ぶ",
  (await page.locator(".cheat-row").count()) >= 25,
  String(await page.locator(".cheat-row").count()));

// 行にカーソルを合わせると、どう置かれるかが動きで出る
const miniState = async () => await page.evaluate(() => ({
  canvas: !!document.querySelector("#cheatDemo canvas"),
  blocks: document.querySelectorAll("#cheatDemo .mb").length,
  shown: document.querySelectorAll("#cheatDemo .mb.in").length,
  note: document.getElementById("cheatNote").textContent,
}));
await page.hover('.cheat-row:has-text("右に続けて置く")');
await page.waitForTimeout(1500);
const mRight = await miniState();
check("カーソルを合わせると見本が出る", mRight.canvas && mRight.blocks === 2, JSON.stringify(mRight));
check("見本に一言の説明がつく",
  /直前のすぐ右へ/.test(mRight.note) && /出力シート/.test(mRight.note), mRight.note);

// 範囲の書き方は「元データ側」で見せる
await page.hover('.cheat-row:has-text("B列からD列")');
await page.waitForTimeout(600);
const mCols = await miniState();
check("範囲の書き方は元データ側で見せる",
  mCols.canvas && /元データ/.test(mCols.note), JSON.stringify(mCols));

// 名前で指す例は、間に別のものを置いても相手のとなりに来ることを見せる
await page.hover('.cheat-row:has-text("見出し の右に置く")');
await page.waitForTimeout(600);
const mNamed = await miniState();
check("名前で指す例は3つ出して違いを見せる", mNamed.blocks === 3, JSON.stringify(mNamed));

// 動いていること（順に増えて、また最初から）
await page.hover('.cheat-row:has-text("繰り返し 行 = 10-14")');
await page.waitForTimeout(250);
const miniSeen = new Set();
for (let i = 0; i < 14; i++) {
  miniSeen.add((await miniState()).shown);
  await page.waitForTimeout(360);
}
check("見本は順に置かれていく（止まった絵ではない）", miniSeen.size >= 4,
  [...miniSeen].sort((a, b) => a - b).join(","));

// 動きのない行（メモなど）は説明だけ
await page.hover('.cheat-row:has-text("説明: 支店ごとの月次行")');
await page.waitForTimeout(400);
const mNone = await miniState();
check("実行されない行には見本を出さない",
  !mNone.canvas && /実行されない/.test(mNone.note), JSON.stringify(mNone));

// 行をクリックすると手順書に貼り付く
await page.fill("#scText", "");
await page.click('.cheat-row:has-text("シート追加 月次まとめ")');
await page.click('.cheat-row:has-text("売上明細のA1:D10を続けて置く")');
check("クリックで手順書に書き足せる",
  (await page.inputValue("#scText")).trim() === "シート追加 月次まとめ\n売上明細のA1:D10を続けて置く",
  JSON.stringify(await page.inputValue("#scText")));

// 一覧に書いてある例が、実際に動くこと（説明と実装が食い違わないように）
const cheatRuns = await page.evaluate(() => {
  const A = window.__app;
  // 「置く」形の行だけを取り出し、名前や繰り返しの前後関係を保って一続きに流す
  const script = [
    "シート追加 検証",
    "売上明細のA1:D10を 検証 のB2に置く",
    "売上明細の3行目を続けて置く",
    "売上明細の3行目から8行目を右に続けて置く",
    "売上明細のB列を続けて置く",
    "売上明細のB列からD列を右に続けて置く",
    "売上明細の全体を続けて置く",
    "売上明細のA1を続けて置く",
    "売上明細のA1:D10を 検証 のB2に置く（転置）",
    "売上明細のA1:C4を 検証 のA1に置く（名前: 目印）",
    "支店別サマリのA1:B4を 目印 の右に置く",
    "商品マスタのA1:C3を 目印 の下に置く",
    "商品マスタのA1:C3を 目印 の下に1行あけて置く",
    "繰り返し 行 = 10-12 / 1列あけて横に並べる",
    "  売上明細の{行}行目を続けて置く",
    "ここまで",
  ].join("\n");
  A.S.out = []; A.S.selBlock = null;
  const r = A.runScript(script, true);
  return { applied: r.applied, errors: r.errors.map((e) => e.n + "行目:" + e.msg) };
});
check("一覧に載せた書き方が実際に動く", cheatRuns.errors.length === 0, cheatRuns.errors.join(" / "));
check("すべての行が置かれる", cheatRuns.applied === 16, String(cheatRuns.applied));

// 矢印の書き方も動く
const arrowRun = await placedByEarly("シート追加 矢印\n売上明細のA1:D10 → 矢印のB2");
check("矢印の書き方も動く", arrowRun.join(",") === "矢印!B2", arrowRun.join(","));

// 繰り返しの欄を開くと、一覧は畳まれる（画面が狭くならないように）
await page.click("#scRepeat");
await page.waitForTimeout(120);
check("繰り返しの欄を開くと一覧は畳まれる",
  !(await page.isVisible("#scCheat")) && (await page.textContent("#scHelp")) === "書き方の一覧");
// 畳んだら動きも止める（見えないところで動かし続けない）
const stopped = await page.evaluate(() => new Promise((r) => {
  const n = () => document.querySelectorAll("#cheatDemo .mb.in").length;
  const a = n();
  setTimeout(() => r(a === n()), 1000);
}));
check("畳むと見本の動きも止まる", stopped);
await page.click("#repeatCancel");
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });

// ---- 5b-2. 名前をつけて、その右・下に置く（案B） -------------------------
/** 手順書を実行して、置かれた先を「出力!位置」で返す */
const placedBy = async (script) => await page.evaluate((sc) => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript(sc, true);
  const col = (c) => { let s = "", n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };
  return A.S.out.flatMap((o) => o.blocks.map((b) => o.name + "!" + col(b.dc) + (b.dr + 1)));
}, script);

const named = await placedBy(`シート追加 まとめ
売上明細のA1:C4をまとめのA1に置く（名前: 見出し）
支店別サマリのA1:B4を 見出し の右に置く
商品マスタのA1:C3を 見出し の下に置く
売上明細のA1:B3を 見出し の右下に置く`);
check("名前をつけた相手の右・下・右下に置ける",
  named.join(",") === "まとめ!A1,まとめ!D1,まとめ!A5,まとめ!D5", named.join(","));
check("名前がブロックに残る",
  (await page.evaluate(() => window.__app.S.out[0].blocks[0].name)) === "見出し",
  await page.evaluate(() => window.__app.S.out[0].blocks[0].name));

const gapped = await placedBy(`シート追加 すきま
売上明細のA1:C4をすきまのA1に置く（名前: 頭）
商品マスタのA1:C3を 頭 の下に1行あけて置く
支店別サマリのA1:B4を 頭 の右に2列あけて置く`);
check("すきまを指定して置ける",
  gapped.join(",") === "すきま!A1,すきま!A6,すきま!F1", gapped.join(","));

// 直前が何であっても、狙った相手のとなりに置ける（ここが「続けて置く」との違い）
const faraway = await placedBy(`シート追加 遠く
売上明細のA1:C2を遠くのA1に置く（名前: 起点）
商品マスタのA1:C2を遠くのA20に置く
支店別サマリのA1:B2を 起点 の右に置く`);
check("直前ではなく、名前で指した相手のとなりに置く",
  faraway.join(",") === "遠く!A1,遠く!A20,遠く!D1", faraway.join(","));

// 名前が無いときは、これまでどおり「直前の右」
const legacy = await placedBy(`シート追加 従来
売上明細のA1:C2を従来のA1に置く
支店別サマリのA1:B2を右に置く`);
check("名前を使わない「右に置く」は今までどおり",
  legacy.join(",") === "従来!A1,従来!D1", legacy.join(","));

// ---- 5b-3. 1 周ぶんを横に並べる（案A） -----------------------------------
const across = await placedBy(`シート追加 横
売上明細のA1:C1を横のA1に置く
繰り返し 行 = 10, 12, 17 / 横に並べる
  売上明細の{行}行目を続けて置く
  支店別サマリの{行}行目を続けて置く
ここまで`);
check("横に並べると、周ごとに右へ移る",
  across.join(",") === "横!A1,横!A2,横!A3,横!F2,横!F3,横!K2,横!K3", across.join(","));

const acrossGap = await placedBy(`シート追加 横すき
繰り返し 行 = 10, 12 / 1列あけて横に並べる
  売上明細の{行}行目を続けて置く
  支店別サマリの{行}行目を続けて置く
ここまで`);
check("横に並べるときも、すきまを指定できる",
  acrossGap.join(",") === "横すき!A1,横すき!A2,横すき!G1,横すき!G2", acrossGap.join(","));

const downGap = await placedBy(`シート追加 縦すき
繰り返し 行 = 10, 12 / 1行あけて縦に積む
  売上明細の{行}行目を続けて置く
ここまで`);
check("縦に積むときも、すきまを指定できる",
  downGap.join(",") === "縦すき!A1,縦すき!A3", downGap.join(","));

// 位置を書いた行も、周ごとに右へついてくる
const acrossAnchor = await placedBy(`シート追加 合わせ
繰り返し 行 = 10, 12, 17 / 横に並べる
  売上明細の{行}行目を合わせのA1に置く
  支店別サマリの{行}行目を続けて置く
ここまで`);
check("位置を書いた行も、周ごとに右へついてくる",
  acrossAnchor.join(",") === "合わせ!A1,合わせ!A2,合わせ!F1,合わせ!F2,合わせ!K1,合わせ!K2",
  acrossAnchor.join(","));

// 何も書かなければ、今までどおり縦に積む
const downDefault = await placedBy(`シート追加 縦
繰り返し 行 = 10, 12, 17
  売上明細の{行}行目を続けて置く
  支店別サマリの{行}行目を右に続けて置く
ここまで`);
check("何も書かなければ、今までどおり縦に積む",
  downDefault.join(",") === "縦!A1,縦!F1,縦!A2,縦!F2,縦!A3,縦!F3", downDefault.join(","));

// 「繰り返しにする」からも横並びを選べる
await page.evaluate(() => window.__app.openScript(true));   // すでに開いていても確実に開く
await page.fill("#scText", "シート追加 ボタン\n売上明細の10行目をボタンのA1に置く");
await page.click("#scRepeat");
await page.waitForTimeout(120);
check("繰り返しの欄で並べ方を選べる", await page.evaluate(() =>
  [...document.querySelectorAll('input[name="repeatDir"]')].map((n) => n.value).join(",")) === "down,across");
await page.click('.repeat-form .rf-kind label:has-text("横に並べる")');
await page.fill("#repeatValues", "10, 12");
await page.click("#repeatOk");
await page.waitForTimeout(120);
check("選んだ並べ方が手順書に入る",
  /繰り返し 行 = 10, 12 \/ 横に並べる/.test(await page.inputValue("#scText")),
  (await page.inputValue("#scText")).split("\n").filter((l) => /繰り返し/.test(l)).join(""));
// 並べ方は次の検証に残るので、既定へ戻しておく（欄は閉じているので直接戻す）
await page.evaluate(() => {
  document.querySelector('input[name="repeatDir"][value="down"]').checked = true;
  window.__app.S.out = []; window.__app.S.selBlock = null;
});

// ---- 5c-2. 列とシートの繰り返し ------------------------------------------
/** 手順書を書いて、種類を選んで、繰り返しにして、実行する */
const makeRepeat = async (script, kind, values) => {
  await page.fill("#scText", script);
  await page.click("#scRepeat");
  await page.waitForTimeout(120);
  await page.click(`.rf-kind label:has-text("${kind}")`);
  await page.waitForTimeout(120);
  const asked = await page.textContent("#repeatLabel");
  const preset = await page.inputValue("#repeatValues");
  if (values != null) await page.fill("#repeatValues", values);
  await page.click("#repeatOk");
  await page.waitForTimeout(120);
  const text = await page.inputValue("#scText");
  await page.click("#scRun");
  await page.waitForTimeout(300);
  const placed = await page.$$eval("#blockList .block-card",
    (ns) => ns.map((n) => n.querySelector(".src").textContent + "→" + n.querySelector(".row2 .to").textContent));
  return { asked, preset, text, placed };
};

// 列：B1:B8 の B が {列} になり、B・C・D の 3 周ぶん置かれる
const byCol = await makeRepeat(
  "シート追加 列まとめ\n売上明細のB1:B8を列まとめのA1に置く", "列", "B, C, D");
check("列の繰り返しでは列名を見つける", /B列のところを、どの列で繰り返しますか/.test(byCol.asked), byCol.asked);
check("見つけた列名が初期値に入る", byCol.preset === "B", byCol.preset);
check("列が {列} に置き換わる",
  /繰り返し 列 = B, C, D\n {2}売上明細の\{列\}1:\{列\}8を列まとめのA1に置く\nここまで/.test(byCol.text),
  byCol.text.split("\n").filter((l) => /繰り返し|\{列\}|ここまで/.test(l)).join(" / "));
check("列ぶん置かれる",
  byCol.placed.join(",") === "売上明細!B1:B8→列まとめ!A1,売上明細!C1:C8→列まとめ!A9,売上明細!D1:D8→列まとめ!A17",
  byCol.placed.join(","));

// 列は B-E のような範囲でも指定できる
const byColRange = await makeRepeat(
  "シート追加 範囲\n売上明細のB1:B4を範囲のA1に置く", "列", "B-D");
check("列は範囲でも指定できる",
  byColRange.placed.join(",") === "売上明細!B1:B4→範囲!A1,売上明細!C1:C4→範囲!A5,売上明細!D1:D4→範囲!A9",
  byColRange.placed.join(","));

// シート：シート名が {シート} になり、開いている 3 シートぶん置かれる
const bySheet = await makeRepeat(
  "シート追加 シートまとめ\n売上明細のA1:C3をシートまとめのA1に置く", "シート", null);
check("シートの繰り返しではシート名を見つける",
  /「売上明細」のところを、どのシートで繰り返しますか/.test(bySheet.asked), bySheet.asked);
check("開いているシート名が初期値に入る",
  bySheet.preset === "売上明細, 支店別サマリ, 商品マスタ", bySheet.preset);
check("シート名が {シート} に置き換わる",
  /繰り返し シート = 売上明細, 支店別サマリ, 商品マスタ\n {2}\{シート\}のA1:C3をシートまとめのA1に置く\nここまで/.test(bySheet.text),
  bySheet.text.split("\n").filter((l) => /繰り返し|\{シート\}|ここまで/.test(l)).join(" / "));
check("シートぶん置かれる",
  bySheet.placed.join(",") === "売上明細!A1:C3→シートまとめ!A1,支店別サマリ!A1:C3→シートまとめ!A4,商品マスタ!A1:C3→シートまとめ!A7",
  bySheet.placed.join(","));

// 手順書の見出しは名前と閉じるだけ
check("手順書の見出しに説明文は出さない",
  !/使い回す|Ctrl/.test(await page.textContent(".modal-head")), await page.textContent(".modal-head"));
check("閉じるは × のボタン", (await page.textContent("#scClose")).trim() === "×",
  await page.textContent("#scClose"));

// 数字を含むシート名を行番号として巻き込まない
const guarded = await page.evaluate(() => {
  const A = window.__app;
  A.S.sheets.push({ name: "支店10", ws: {}, rows: 1, cols: 1, usedRows: 1, usedCols: 1, merges: [], colsMeta: [] });
  const r = A.commonRowNumber(["支店10のA10:E10"]);
  A.S.sheets.pop();
  return r;
});
check("シート名の数字も候補には入る（保護は置換時）", guarded === "10", String(guarded));

// 別の人に渡す想定でファイルに書き出す（ブラウザには保存しない）
check("ブラウザ保存のボタンは無い",
  (await page.locator("#scSave").count()) === 0 && (await page.locator("#scList").count()) === 0
  && (await page.locator("#scName").count()) === 0);
const onScreen = await page.inputValue("#scText");
const [dl3] = await Promise.all([page.waitForEvent("download"), page.click("#scExport")]);
check("手順書をファイルに書き出せる", dl3.suggestedFilename() === "サンプル売上_手順書.txt", dl3.suggestedFilename());
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
const scErrToast = await page.textContent("#toasts");
check("無いシートは行番号つきで報告される",
  /1行目.*見つかりません/.test(scErrToast), scErrToast.slice(0, 70));
check("ログ欄そのものが無い", (await page.locator("#log").count()) === 0);

await page.click("#scClose");                       // ここから先は本体画面をさわる

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
await openOnly(bigPath, "大きい売上.xlsx", 30000);
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
await loadSampleAgain();
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9d. Excel で開いたときと同じ見え方（色・非表示・幅・固定） ----------
await openOnly(join(root, "test/fixtures/書式つき.xlsx"), "書式つき.xlsx");

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
check("非表示の数がチェックボックスに出る",
  /非表示（1行\/1列）も表示/.test(await page.textContent("#hiddenNote")), await page.textContent("#hiddenNote"));
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
await page.waitForTimeout(350);   // 固定中は区切りに合わせ直すので、落ち着いてから測る
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
await page.waitForTimeout(350);
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
  /A2/.test(await selInfo()), await selInfo());
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
check("移動しても目印は出さない", (await page.locator("#dstGrid .dropghost").count()) === 0);
check("空のときの移動先は A1", (await page.evaluate(() => window.__app.S.lastGoto)) === "A1", await page.evaluate(() => window.__app.S.lastGoto));

// 移動先のセルへドラッグして置く（選択範囲は掴んだままなので、そのまま運べる）
await dragSelectionToRef(await page.evaluate(() => window.__app.S.lastGoto));
check("選択範囲の中を掴んでドラッグできる",
  (await page.locator("#blockList .block-card").count()) === 1,
  String(await page.locator("#blockList .block-card").count()));
check("移動先の位置に置かれる",
  (await page.textContent("#blockList .block-card .row2 .to")) === "抜粋1!A1",
  await page.textContent("#blockList .block-card .row2 .to"));

// 末尾の右へ移動 → その位置へドラッグ
await selectRange("A3", "F3");
await page.click("#btnAppendRight");
check("末尾の右は直前のブロックの右隣", (await page.evaluate(() => window.__app.S.lastGoto)) === "G1", await page.evaluate(() => window.__app.S.lastGoto));
await dragSelectionToRef(await page.evaluate(() => window.__app.S.lastGoto));
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
const collapsed = await selInfo();
check("選択の中をクリックすると1セルになる", /A3:A3|A3\b/.test(collapsed) && /1×1/.test(collapsed), collapsed);

// 1行目に置いたときも、ラベルが列見出しに隠れず全部読めること
await page.evaluate(() => window.__app.runScript("シート追加 先頭\n書式つきのA2:F2を先頭のA1に置く", true));
await clickBlock();
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
await clickBlock();
const midTag = await page.evaluate(() => {
  const tag = document.querySelector("#dstGrid .blockbox .tag");
  const grid = document.getElementById("dstGrid").getBoundingClientRect();
  const t = tag.getBoundingClientRect();
  return { below: tag.classList.contains("below"), fullyVisible: t.top >= grid.top - 0.5,
    st: document.getElementById("dstGrid").scrollTop };
});
check("余裕があればラベルは上のまま", !midTag.below && midTag.fullyVisible, JSON.stringify(midTag));

// 元データ側の選択には、範囲を書いたラベル（つまみ）を出さない
await selectRange("A1", "C1");
check("元データの選択にラベルを出さない",
  (await page.locator("#srcGrid .selbox .grab").count()) === 0);
check("選択の枠そのものは掴んで運べる",
  (await page.getAttribute("#srcGrid .selbox", "draggable")) === "true");
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
// 移動先のセルが画面内の真ん中あたりに来ること（目印は出さないので位置で確かめる）
const moved = await page.evaluate(() => {
  const w = document.getElementById("dstGrid"), G = window.__app.dstGrid;
  const p = window.__app.parseCell(window.__app.S.lastGoto);
  const box = G.box({ r1: p.r, c1: p.c, r2: p.r, c2: p.c });   // キャンバス内の位置
  const view = { top: w.scrollTop, height: w.clientHeight };
  const centerY = box.top + box.height / 2 - view.top;
  return { scrollTop: w.scrollTop, inView: centerY > 0 && centerY < view.height,
    offset: Math.abs(centerY - view.height / 2), viewH: view.height };
});
check("移動先が画面外ならそこまで表示が動く", moved.scrollTop > 200, `scrollTop=${Math.round(moved.scrollTop)}`);
check("移動先のセルが画面内に入る", moved.inView, JSON.stringify(moved));
check("移動先は画面の真ん中あたりに来る", moved.offset < moved.viewH * 0.25,
  `中心から ${Math.round(moved.offset)}px（画面高 ${Math.round(moved.viewH)}px）`);
check("移動先は末尾の下 A61", (await page.evaluate(() => window.__app.S.lastGoto)) === "A61", await page.evaluate(() => window.__app.S.lastGoto));

// 文章で指示したときは、置いたブロックまで表示が動いて光る
await page.evaluate(() => window.__app.runScript("書式つきのA2:F2を遠いのA80に置く", false));
await page.waitForTimeout(200);
check("文章で置いたときも表示が追いかける",
  (await page.evaluate(() => document.getElementById("dstGrid").scrollTop)) > 1200,
  String(Math.round(await page.evaluate(() => document.getElementById("dstGrid").scrollTop))));
check("置いた直後は光って知らせる",
  (await page.locator("#dstGrid .blockbox.flash").count()) === 1);

// あとの検証のためサンプルに戻す
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await loadSampleAgain();
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9e-2. 元データを複数のファイルで開く -------------------------------
await loadSampleAgain();
const sheetsBefore = await page.evaluate(() => window.__app.S.sheets.length);
await page.setInputFiles("#fileInput", join(root, "test/fixtures/書式つき.xlsx"));
await page.waitForFunction(() => window.__app.S.files.length === 2, { timeout: 20000 });
const two = await page.evaluate(() => ({
  files: window.__app.S.files.map((f) => f.name),
  sheets: window.__app.S.sheets.map((sh) => sh.name + "@" + sh.file),
  active: window.__app.S.active,
  chip: document.getElementById("fileName").textContent,
  meta: document.getElementById("fileMeta").textContent,
  heads: [...document.querySelectorAll("#sheetList .file-head .fn")].map((n) => n.textContent),
}));
check("2つめのファイルを開いても1つめは残る", two.files.length === 2, JSON.stringify(two.files));
check("シートは足されていく",
  two.sheets.length === sheetsBefore + 2, JSON.stringify(two.sheets));
check("足したファイルの先頭シートを見せる", two.active === sheetsBefore, String(two.active));
check("上の表示はファイル数とシート数になる",
  /ほか1件/.test(two.chip) && /2ファイル/.test(two.meta), two.chip + " / " + two.meta);
check("シート一覧はファイルごとにまとまる",
  two.heads.join(",") === "サンプル売上.xlsx,書式つき.xlsx", two.heads.join(","));

// 元データの見出しでファイルを切り替えられる
const fileTabs = await page.$$eval("#srcFiles .filetab .fnm", (ns) => ns.map((n) => n.textContent));
check("元データの見出しにファイルが並ぶ",
  fileTabs.join(",") === "サンプル売上.xlsx,書式つき.xlsx", fileTabs.join(","));
check("いま見ているファイルが分かる", await page.evaluate(() =>
  document.querySelectorAll('#srcFiles .filetab[aria-current="true"] .fnm')[0].textContent) === "書式つき.xlsx");
check("シートのタブは、そのファイルのぶんだけ", await page.evaluate(() =>
  [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")) === "書式つき,通常",
  await page.evaluate(() => [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")));
// 1つめのファイルへ戻る
await page.click('#srcFiles .filetab:has-text("サンプル売上.xlsx")');
await page.waitForTimeout(200);
check("ファイルをクリックすると切り替わる", await page.evaluate(() =>
  [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")) === "売上明細,支店別サマリ,商品マスタ",
  await page.evaluate(() => [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")));
check("切り替えると中身も入れ替わる",
  (await page.textContent('#srcGrid .gc[data-r="0"][data-c="0"]')) === "売上明細",
  await page.textContent('#srcGrid .gc[data-r="0"][data-c="0"]'));
// 見ていたシートを覚えていて、戻ると同じところを見せる
await page.click("#srcTabs .tab >> nth=2");
await page.waitForTimeout(150);
await page.click('#srcFiles .filetab:has-text("書式つき.xlsx")');
await page.waitForTimeout(150);
await page.click('#srcFiles .filetab:has-text("サンプル売上.xlsx")');
await page.waitForTimeout(200);
check("戻ってくると前に見ていたシートを開く", await page.evaluate(() =>
  window.__app.S.sheets[window.__app.S.active].name) === "商品マスタ",
  await page.evaluate(() => window.__app.S.sheets[window.__app.S.active].name));
// 左のシート一覧から別のファイルのシートを選ぶと、ファイルごと切り替わる
await page.click('#sheetList .sheet-item:has-text("通常")');
await page.waitForTimeout(200);
check("シート一覧から選ぶとファイルも切り替わる", await page.evaluate(() =>
  window.__app.S.files[window.__app.S.activeFile].name) === "書式つき.xlsx",
  await page.evaluate(() => window.__app.S.files[window.__app.S.activeFile].name));
check("2ファイル以上のときだけファイルの並びを出す", await page.isVisible("#srcFiles"));
await page.click('#srcFiles .filetab:has-text("サンプル売上.xlsx")');
await page.waitForTimeout(150);
// 片方を閉じて 1 ファイルに戻すと、並びは引っ込む
await page.evaluate(() => window.__app.removeFile(1));
await page.waitForTimeout(200);
check("1ファイルのときは並びを出さない",
  !(await page.isVisible("#srcFiles")) && (await page.evaluate(() => window.__app.S.files.length)) === 1);
// もう一度開いて、続きのテストに備える
await page.setInputFiles("#fileInput", join(root, "test/fixtures/書式つき.xlsx"));
await page.waitForFunction(() => window.__app.S.files.length === 2, { timeout: 20000 });

// 別のファイルのシートからも、いつもどおり置ける
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript("シート追加 合体\n売上明細のA1:C3を合体のA1に置く\n書式つきのA1:C3を合体のA5に置く", true);
});
await page.waitForTimeout(200);
const mixed = await page.$$eval("#blockList .block-card",
  (ns) => ns.map((n) => n.querySelector(".src").textContent));
check("別々のファイルのシートを 1 枚にまとめられる",
  mixed.join(",") === "売上明細!A1:C3,書式つき!A1:C3", mixed.join(","));

// 生成すると、両方のファイルの中身が 1 つのブックに入る
const [dlMix] = await Promise.all([page.waitForEvent("download"), page.click("#btnGen")]);
const mixPath = join(tmp, "mixed.xlsx");
await dlMix.saveAs(mixPath);
const mixWb = XLSX.read(readFileSync(mixPath), { type: "buffer" });
const mixWs = mixWb.Sheets["合体"];
check("1つめのファイルの値が入る", mixWs["A1"] && mixWs["A1"].v === "売上明細", mixWs["A1"] && String(mixWs["A1"].v));
check("2つめのファイルの値も入る",
  mixWs["A5"] && mixWs["A5"].v === "月次売上レポート", mixWs["A5"] && String(mixWs["A5"].v));
check("書き出す名前は1つめのファイルから",
  dlMix.suggestedFilename() === "サンプル売上_抜粋.xlsx", dlMix.suggestedFilename());

// 手順書には開いているファイルが全部載る
await page.click("#btnScript");
await page.click("#scFromBlocks");
const multiScript = await page.inputValue("#scText");
check("手順書に元ファイルが全部載る",
  /# 元ファイル: サンプル売上\.xlsx, 書式つき\.xlsx/.test(multiScript),
  multiScript.split("\n")[0]);
await page.click("#scClose");

// ファイルを閉じると、そのファイルを使っていたブロックも一緒に消える
await page.evaluate(() => window.__app.removeFile(1));
await page.waitForTimeout(200);
const afterClose = await page.evaluate(() => ({
  files: window.__app.S.files.map((f) => f.name),
  sheets: window.__app.S.sheets.map((sh) => sh.name),
  blocks: window.__app.S.out.flatMap((o) => o.blocks.map((b) => window.__app.S.sheets[b.sheet].name)),
}));
check("ファイルを閉じられる", afterClose.files.join(",") === "サンプル売上.xlsx", afterClose.files.join(","));
check("閉じたファイルのシートも消える",
  afterClose.sheets.join(",") === "売上明細,支店別サマリ,商品マスタ", afterClose.sheets.join(","));
check("閉じたファイルのブロックだけが外れる",
  afterClose.blocks.join(",") === "売上明細", afterClose.blocks.join(","));

// 同じ名前のシートがあるときは、ファイル名で指定できる
const dupPath = join(tmp, "もう一つ.xlsx");
{
  const w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([["別の売上明細"], ["x", "y"]]), "売上明細");
  writeFileSync(dupPath, XLSX.write(w, { bookType: "xlsx", type: "buffer" }));
}
await page.setInputFiles("#fileInput", dupPath);
await page.waitForFunction(() => window.__app.S.files.length === 2, { timeout: 20000 });
// シートのタブは、いま見ているファイルのぶんだけ
check("タブはいま見ているファイルのシートだけ", await page.evaluate(() =>
  [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")) === "売上明細",
  await page.evaluate(() => [...document.querySelectorAll("#srcTabs .tab")].map((n) => n.textContent).join(",")));
const picked = await page.evaluate(() => ({
  plain: window.__app.matchSheetName("売上明細のA1:B2"),
  qualified: window.__app.matchSheetName("[もう一つ.xlsx]売上明細のA1:B2"),
}));
check("ファイル名で指定すると、そのファイルのシートになる",
  picked.qualified.i === 3 && picked.plain.i === 0, JSON.stringify(picked));
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.runScript("シート追加 同名\n[もう一つ.xlsx]売上明細のA1:A1を同名のA1に置く", true);
});
await page.waitForTimeout(200);
check("手順書でもファイル名で指定できる", await page.evaluate(() => {
  const b = window.__app.S.out[0].blocks[0];
  return b && window.__app.S.sheets[b.sheet].file === 1;
}));
// 同名があるときは、今の配置から作る手順書もファイル名つきになる
await page.click("#btnScript");
await page.click("#scFromBlocks");
check("同名シートは手順書でもファイル名つきで書く",
  /\[もう一つ\.xlsx\]売上明細の/.test(await page.inputValue("#scText")),
  (await page.inputValue("#scText")).split("\n").filter((l) => /置く/.test(l)).join(" / "));
await page.click("#scClose");

await loadSampleAgain();
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });

// ---- 9f. 末尾は「値の入っている範囲」で判断する -------------------------
// 行まるごとの選択は使用範囲いっぱい（末尾は空欄だらけ）になる。
// 空欄まで末尾に数えると、ずっと右の何も無い場所へ飛んでしまう
await openOnly(join(root, "test/fixtures/横長.xlsx"), "横長.xlsx");
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
check("末尾の右は値の右隣（空欄は数えない）", (await page.evaluate(() => window.__app.S.lastGoto)) === "D1", await page.evaluate(() => window.__app.S.lastGoto));
check("遠くまで飛ばない",
  (await page.evaluate(() => document.getElementById("dstGrid").scrollLeft)) < 300,
  String(Math.round(await page.evaluate(() => document.getElementById("dstGrid").scrollLeft))));
await page.click("#btnAppendDown");
check("末尾の下も値の下（空欄は数えない）", (await page.evaluate(() => window.__app.S.lastGoto)) === "A2", await page.evaluate(() => window.__app.S.lastGoto));

// 本当に全列に値がある行なら、これまでどおり右端の続きへ
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null;
  A.setSelection(0, 0, 0, 79);
  A.addBlock({ sheet: 0, r1: 0, c1: 0, r2: 0, c2: 79 }, 0, 0, A.S.out[0] || A.S.out[A.S.activeOut]);
});
await page.click("#btnAppendRight");
check("値が全列にある行では右端の続きへ行く", (await page.evaluate(() => window.__app.S.lastGoto)) === "CC1", await page.evaluate(() => window.__app.S.lastGoto));

await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await loadSampleAgain();
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9f-2. 横に重ねて 120 列を超えても壊れない -------------------------
// 描画列数に上限があると、その先の位置が座標を持たず NaN になり、
// 目印も表示も左上（A1）へ落ちてしまっていた
await openOnly(join(root, "test/fixtures/横長.xlsx"), "横長.xlsx");
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
const wide = [];
for (let i = 0; i < 5; i++) {
  await page.evaluate((n) => {
    const A = window.__app;
    const o = A.S.out[A.S.activeOut] || (A.S.out.push({ id: "ow", name: "横並び", blocks: [] }), A.S.out[0]);
    A.setSelection(1, 0, 1, 29, 0);                       // 30列ぶんの行
    A.addBlock({ sheet: 0, r1: 1, c1: 0, r2: 1, c2: 29 }, 0, 1 + n * 30, o);
  }, i);
  await page.click("#btnAppendRight");
  wide.push(await page.evaluate(() => {
    const w = document.getElementById("dstGrid"), A = window.__app;
    const p = A.parseCell(A.S.lastGoto);
    const box = A.dstGrid.box({ r1: p.r, c1: p.c, r2: p.r, c2: p.c });
    return { label: A.S.lastGoto, left: box.left, scrollLeft: Math.round(w.scrollLeft) };
  }));
}
check("120列を超えても位置が計算できる（NaNにならない）",
  wide.every((x) => Number.isFinite(x.left) && x.left > 0), JSON.stringify(wide.map((x) => x.left)));
check("120列を超えても左上へ戻らない",
  wide.every((x) => x.scrollLeft > 1000), JSON.stringify(wide.map((x) => x.scrollLeft)));
check("末尾は右へ伸び続ける",
  wide.map((x) => x.label).join(",") === "AF1,BJ1,CN1,DR1,EV1",
  wide.map((x) => x.label).join(","));
check("120列を超えても左上へ戻らない（再掲）", wide[4].scrollLeft > 10000, String(wide[4].scrollLeft));

// ---- 9f-3. 固定した見出しはスクロールしても中身から離れない --------------
// 貼り付ける処理を描き直しの中に置いていたころは、同じ範囲を描いたままの細かい
// スクロールで見出しだけが取り残され、固定した列の中身と列名が横にずれていった
await page.evaluate(() => { window.__app.S.out = []; window.__app.S.selBlock = null; });
await page.evaluate(() => window.__app.openHeaderMenu("c", 3, 300, 200));
await page.click('.hmenu button:has-text("この列の左で固定")');
await page.evaluate(() => { document.getElementById("srcGrid").scrollLeft = 900; });
await page.waitForTimeout(200);

/** 固定した列の「列名」と「中身」の左端のずれ、および貼り付き位置 */
const frozenState = async () => await page.evaluate(() => {
  const w = document.getElementById("srcGrid"), wr = w.getBoundingClientRect();
  const q = (sel) => [...document.querySelectorAll(sel)]
    .map((n) => ({ c: n.dataset.c, x: Math.round((n.getBoundingClientRect().x - wr.x) * 10) / 10 }))
    .filter((o) => o.c !== undefined);
  const heads = q("#srcGrid .gfrzhead .gh"), cells = q('#srcGrid .gfrz-left .gc[data-r="0"]');
  return {
    sl: Math.round(w.scrollLeft),
    band: Math.round((document.querySelector("#srcGrid .gfrz-left").getBoundingClientRect().x - wr.x) * 10) / 10,
    gaps: heads.map((h) => {
      const cell = cells.find((c) => c.c === h.c);
      return cell ? Math.round((h.x - cell.x) * 10) / 10 : null;
    }),
  };
});
const drifts = [];
for (const px of [4, 9, 3, 7, 120, -60]) {
  await page.evaluate((d) => { document.getElementById("srcGrid").scrollLeft += d; }, px);
  await page.waitForTimeout(120);
  drifts.push(await frozenState());
}
check("固定した列名は中身の真上から動かない",
  drifts.every((d) => d.gaps.length > 0 && d.gaps.every((g) => g === 0)),
  JSON.stringify(drifts.map((d) => d.gaps)));
check("固定した列そのものも左端に貼り付いたまま",
  drifts.every((d) => d.band === 0), JSON.stringify(drifts.map((d) => d.band)));
// 貼り付けはブラウザ任せ（position:sticky）にしておく。JS でスクロール量ぶん
// ずらすやり方に戻すと、1 フレーム遅れや小数のずれで固定した部分が細かく震える
const pinning = await page.evaluate(() => {
  const pins = [...document.querySelectorAll("#srcGrid .gpin")];
  const layers = ["gcolh", "growh", "gcorner", "gfrz-top", "gfrz-left", "gfrz-corner", "gfrzhead"];
  return {
    n: pins.length,
    allSticky: pins.every((n) => getComputedStyle(n).position === "sticky"),
    noTransform: layers.every((c) => {
      const n = document.querySelector("#srcGrid ." + c);
      if (!n) return true;
      const t = getComputedStyle(n).transform;
      return t === "none" || t === "matrix(1, 0, 0, 1, 0, 0)";
    }),
  };
});
check("貼り付けはブラウザ任せ（sticky）", pinning.n === 7 && pinning.allSticky, JSON.stringify(pinning));
check("層を JS でずらしてはいない", pinning.noTransform, JSON.stringify(pinning));
// スクロールの位置には手を出さない（吸い付かせるとブラウザ側の動きと引っぱり合って震える）
check("動かした量がそのまま残る（吸い付かない）",
  drifts.map((d) => d.sl).join(",") === "904,913,916,923,1043,983",
  drifts.map((d) => d.sl).join(","));
await page.screenshot({
  path: join(root, "test/shots/freeze-scrolled.png"),
  clip: { x: 233, y: 90, width: 1000, height: 420 },
});

// 行を固定したときは縦も同じ
await page.evaluate(() => {
  const w = document.getElementById("srcGrid");
  w.scrollLeft = 0; w.scrollTop = 0;
  window.__app.S.sheets[window.__app.S.active].freeze = null;
});
await page.evaluate(() => window.__app.openHeaderMenu("r", 2, 300, 200));
await page.click('.hmenu button:has-text("この行の上で固定")');
const rowDrifts = [];
for (const px of [137, 5, 11, 40]) {
  await page.evaluate((d) => { document.getElementById("srcGrid").scrollTop += d; }, px);
  await page.waitForTimeout(120);
  rowDrifts.push(await page.evaluate(() => {
    const w = document.getElementById("srcGrid"), wr = w.getBoundingClientRect();
    const q = (sel) => [...document.querySelectorAll(sel)]
      .map((n) => ({ r: n.dataset.r, y: Math.round((n.getBoundingClientRect().y - wr.y) * 10) / 10 }))
      .filter((o) => o.r !== undefined);
    const heads = q("#srcGrid .gfrzhead .gh"), cells = q('#srcGrid .gfrz-top .gc[data-c="0"]');
    return { st: Math.round(w.scrollTop), gaps: heads.map((h) => {
      const cell = cells.find((c) => c.r === h.r);
      return cell ? Math.round((h.y - cell.y) * 10) / 10 : null;
    }) };
  }));
}
check("固定した行番号も中身の真横から動かない",
  rowDrifts.every((d) => d.gaps.length > 0 && d.gaps.every((g) => g === 0)),
  JSON.stringify(rowDrifts.map((d) => d.gaps)));
check("縦も動かした量がそのまま残る",
  rowDrifts.map((d) => d.st).join(",") === "137,142,153,193",
  rowDrifts.map((d) => d.st).join(","));

await page.evaluate(() => {
  const w = document.getElementById("srcGrid");
  window.__app.S.sheets[window.__app.S.active].freeze = null;
  window.__app.S.out = []; window.__app.S.selBlock = null; window.__app.S.sel = null;
  w.scrollTop = 0; w.scrollLeft = 0;
});
await loadSampleAgain();
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9f-3. 出力シートで範囲を選ぶと集計が出る ---------------------------
// 支店別サマリ A2:C7（見出し + 4支店 + 合計）を置いて、数値列の集計を確かめる
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = []; A.S.selBlock = null; A.S.dsel = null;
  A.runScript("シート追加 集計確認\n支店別サマリのA2:C7を集計確認のA1に置く", true);
});
await page.waitForTimeout(150);
check("選ぶ前は集計を出さない", !(await page.isVisible("#dstStats")));

// C列（売上合計）を列見出しのクリックで選ぶ
const colC = await page.locator('#dstGrid .gcolh .gh[data-c="2"]').boundingBox();
await page.mouse.click(colC.x + 20, colC.y + 10);
await page.waitForTimeout(120);
check("列を選ぶと集計が出る", await page.isVisible("#dstStats"));
check("集計はシートタブと同じ帯に出る", await page.evaluate(() => {
  const f = document.getElementById("dstFoot"), st = document.getElementById("dstStats");
  const tab = document.querySelector("#dstTabsHost .tab");
  return f.contains(st) && f.contains(tab)
    && Math.abs(st.getBoundingClientRect().y - tab.getBoundingClientRect().y) < 1;
}));
const footH0 = await page.evaluate(() => Math.round(document.getElementById("dstFoot").getBoundingClientRect().height));
const stats = await page.textContent("#dstStats");
// 支店別サマリ: 見出し「売上合計」+ 375600 / 432700 / 558800 / 569300 / 合計1936400
check("データの個数が出る", /個数\s*6/.test(stats), stats);
check("数値の個数が出る", /数値\s*5/.test(stats), stats);
check("合計が出る", /合計\s*3,872,800/.test(stats), stats);
check("平均が出る", /平均\s*774,560/.test(stats), stats);
check("最大が出る", /最大\s*1,936,400/.test(stats), stats);
check("最小が出る", /最小\s*375,600/.test(stats), stats);

check("集計が出ても帯の高さは変わらない（グリッドがずれない）", await page.evaluate((h) => {
  const f = document.getElementById("dstFoot");
  return Math.round(f.getBoundingClientRect().height) === h;
}, footH0), String(footH0));

// 数値のない範囲では個数だけ
await page.evaluate(() => window.__app.setDstSel(0, 0, 0, 2));
check("数値が無ければ個数だけ出す",
  /個数\s*3/.test(await page.textContent("#dstStats"))
  && !/合計/.test(await page.textContent("#dstStats")),
  await page.textContent("#dstStats"));

// データの上でも Shift+ドラッグなら選べる（そのままのドラッグはブロックの移動が優先）
const c1 = await page.locator('#dstGrid .gc[data-r="1"][data-c="1"]').boundingBox();
const c2 = await page.locator('#dstGrid .gc[data-r="4"][data-c="2"]').boundingBox();
await page.keyboard.down("Shift");
await page.mouse.move(c1.x + 40, c1.y + 12);
await page.mouse.down();
await page.mouse.move(c2.x + 40, c2.y + 12, { steps: 8 });
await page.mouse.up();
await page.keyboard.up("Shift");
check("Shift+ドラッグで範囲を選べる",
  /^B2:C5/.test(await page.textContent("#dstStats")), await page.textContent("#dstStats"));
check("選んだ範囲の合計が出る",
  /数値\s*8/.test(await page.textContent("#dstStats")), await page.textContent("#dstStats"));
// そのままのドラッグではブロックが動く（選択にはならない）
const dcBefore = await page.evaluate(() => window.__app.S.out[0].blocks[0].dc);
await page.mouse.move(c1.x + 40, c1.y + 12);
await page.mouse.down();
await page.mouse.move(c2.x + 40, c2.y + 12, { steps: 8 });
await page.mouse.up();
check("そのままのドラッグはブロックの移動になる",
  (await page.evaluate(() => window.__app.S.out[0].blocks[0].dc)) !== dcBefore,
  `dc ${dcBefore} → ${await page.evaluate(() => window.__app.S.out[0].blocks[0].dc)}`);

// ブロックをクリックすると、そのブロックぶんの集計になる
await page.evaluate(() => {
  const A = window.__app;
  A.clearDstSel();
  A.runScript("シート追加 集計確認\n支店別サマリのA2:C7を集計確認のA1に置く", true);
});
await page.waitForTimeout(150);
await page.click("#dstGrid .blockbox");
await page.waitForTimeout(120);
check("ブロックを選ぶとその範囲の集計になる",
  /^A1:C6/.test(await page.textContent("#dstStats")), await page.textContent("#dstStats"));

await page.evaluate(() => { window.__app.clearDstSel(); window.__app.S.out = []; window.__app.S.selBlock = null; });
await loadSampleAgain();
await page.waitForFunction(() => window.__app.S.fileName === "サンプル売上.xlsx");

// ---- 9g. 画面まわり（折りたたみ・見出し・+ボタン・並べ方） ---------------
// ヘッダは名前だけ
const heads = await page.$$eval(".pane-head", (ns) => ns.map((n) => n.textContent.trim()));
check("ペインの見出しは名前と操作だけになる",
  /^‹\s*ブック構成$/.test(heads[0]) && /^元データ\s*⤢?/.test(heads[1]) && /^出力シート\s*⤢?/.test(heads[2]),
  heads.join(" | "));
check("選択範囲の表示は出さない", (await page.locator("#selInfo").count()) === 0);

// 左レールは項目ごとに折りたためる
check("最初は開いている", (await page.getAttribute("#foldSheets", "aria-expanded")) === "true");
check("中身が見えている", await page.isVisible("#sheetList"));
await page.click("#foldSheets");
check("折りたたむと中身が隠れる",
  !(await page.isVisible("#sheetList")) && (await page.getAttribute("#foldSheets", "aria-expanded")) === "false");
check("項目名は残る", await page.isVisible("#foldSheets"));
await page.click("#foldBlocks");
check("抜粋ブロックも折りたためる", !(await page.isVisible("#blockList")));
await page.click("#foldSheets");
await page.click("#foldBlocks");
check("開き直せる", (await page.isVisible("#sheetList")) && (await page.isVisible("#blockList")));

// シートを増やす + ボタン（タブの右隣）
const tabsBefore = await page.locator("#dstTabsHost .tab").count();
await page.click("#dstTabsHost .tab-add");
check("タブの＋で出力シートが増える",
  (await page.locator("#dstTabsHost .tab").count()) === tabsBefore + 1,
  `${tabsBefore} → ${await page.locator("#dstTabsHost .tab").count()}`);
check("＋はタブの右端にある",
  await page.evaluate(() => {
    const host = document.getElementById("dstTabsHost");
    return host.lastElementChild.classList.contains("tab-add");
  }));

// シートのタブは横にだけ送れればよい。overflow-x だけ指定すると縦が auto になり、
// 1px 溢れただけで上下のスクロールバーが出てしまう
const tabScroll = async () => await page.evaluate(() =>
  ["srcTabs", "dstTabsHost"].map((id) => {
    const n = document.getElementById(id), cs = getComputedStyle(n);
    return { id, y: cs.overflowY, x: cs.overflowX,
      overflowsY: n.scrollHeight > n.clientHeight, canScrollX: n.scrollWidth > n.clientWidth };
  }));
const tabs0 = await tabScroll();
check("シートのタブに上下のスクロールバーを出さない",
  tabs0.every((t) => t.y === "hidden" && !t.overflowsY), JSON.stringify(tabs0));
check("横は送れるままにする", tabs0.every((t) => t.x === "auto"), JSON.stringify(tabs0));
// タブが並びきらないときは、実際に横へ送れること
for (let i = 0; i < 8; i++) await page.click("#dstTabsHost .tab-add");
await page.waitForTimeout(150);
const tabsMany = await page.evaluate(() => {
  const n = document.getElementById("dstTabsHost");
  n.scrollLeft = 9999;
  return { canScrollX: n.scrollWidth > n.clientWidth, moved: Math.round(n.scrollLeft),
    overflowsY: n.scrollHeight > n.clientHeight };
});
check("タブが並びきらないときは横へ送れる",
  tabsMany.canScrollX && tabsMany.moved > 0, JSON.stringify(tabsMany));
check("タブが増えても上下には溢れない", !tabsMany.overflowsY, JSON.stringify(tabsMany));
// 増やしたぶんは片づける
await page.evaluate(() => {
  const A = window.__app;
  A.S.out = A.S.out.slice(0, 2); A.S.activeOut = 0; A.S.selBlock = null;
});
await page.click("#dstTabsHost .tab >> nth=0");
await page.waitForTimeout(120);

// ブック構成そのものも畳める（元データ・出力シートを広く使うため）
const railWide = await page.evaluate(() => document.querySelector(".pane-rail").getBoundingClientRect().width);
const srcWide = await page.evaluate(() => document.querySelector(".pane-src").getBoundingClientRect().width);
await page.click("#foldRail");
await page.waitForTimeout(250);
const railNarrow = await page.evaluate(() => document.querySelector(".pane-rail").getBoundingClientRect().width);
const srcNarrow = await page.evaluate(() => document.querySelector(".pane-src").getBoundingClientRect().width);
check("ブック構成を畳むと細くなる", railNarrow < 40 && railNarrow < railWide, `${Math.round(railWide)} → ${Math.round(railNarrow)}`);
check("畳んだぶん元データが広がる", srcNarrow > srcWide + 50, `${Math.round(srcWide)} → ${Math.round(srcNarrow)}`);
check("畳んでも開くボタンは残る", await page.isVisible("#foldRail"));
check("畳むと中身は隠れる", !(await page.isVisible("#sheetList")));
check("畳んでもグリッドは描かれている",
  (await page.$$eval("#srcGrid .gc", (ns) => ns.length)) > 20,
  String(await page.$$eval("#srcGrid .gc", (ns) => ns.length)));
await page.click("#foldRail");
await page.waitForTimeout(250);
check("開き直せる",
  (await page.isVisible("#sheetList"))
  && (await page.evaluate(() => document.querySelector(".pane-rail").getBoundingClientRect().width)) > 150);

// 元データと出力シートの境目をドラッグして広さを変える
const beforeSplit = await page.evaluate(() => ({
  src: Math.round(document.querySelector(".pane-src").getBoundingClientRect().width),
  dst: Math.round(document.querySelector(".pane-dst").getBoundingClientRect().width),
}));
const sp = await page.locator("#splitter").boundingBox();
await page.mouse.move(sp.x + sp.width / 2, sp.y + sp.height / 2);
await page.mouse.down();
await page.mouse.move(sp.x + 220, sp.y + sp.height / 2, { steps: 10 });
await page.mouse.up();
const afterSplit = await page.evaluate(() => ({
  src: Math.round(document.querySelector(".pane-src").getBoundingClientRect().width),
  dst: Math.round(document.querySelector(".pane-dst").getBoundingClientRect().width),
}));
check("境目を右へ動かすと元データが広がる", afterSplit.src > beforeSplit.src + 150,
  `${beforeSplit.src} → ${afterSplit.src}`);
check("そのぶん出力シートは狭くなる", afterSplit.dst < beforeSplit.dst - 150,
  `${beforeSplit.dst} → ${afterSplit.dst}`);
check("動かしてもセルは描かれている",
  (await page.$$eval("#dstGrid .gc", (ns) => ns.length)) > 10,
  String(await page.$$eval("#dstGrid .gc", (ns) => ns.length)));
// 左へ戻すと出力シートが広がる
await page.mouse.move(sp.x + 220, sp.y + sp.height / 2);
await page.mouse.down();
await page.mouse.move(sp.x - 200, sp.y + sp.height / 2, { steps: 10 });
await page.mouse.up();
const leftSplit = await page.evaluate(() => Math.round(document.querySelector(".pane-dst").getBoundingClientRect().width));
check("左へ動かすと出力シートが広がる", leftSplit > afterSplit.dst + 300, `${afterSplit.dst} → ${leftSplit}`);
// ダブルクリックで半々に戻る
await page.dblclick("#splitter");
await page.waitForTimeout(150);
const evened = await page.evaluate(() => ({
  src: Math.round(document.querySelector(".pane-src").getBoundingClientRect().width),
  dst: Math.round(document.querySelector(".pane-dst").getBoundingClientRect().width),
}));
check("ダブルクリックで半分に戻る", Math.abs(evened.src - evened.dst) < 10,
  `${evened.src} / ${evened.dst}`);

// ---- 表示倍率（元データ・出力シートそれぞれ） ----------------------------
// 固定が残っていると同じセルが 2 つの層に出るので、ここでは外しておく
await page.evaluate(() => {
  const A = window.__app;
  A.S.sheets[A.S.active].freeze = null; A.S.sel = null;
  document.getElementById("srcGrid").scrollTop = 0;
  document.getElementById("srcGrid").scrollLeft = 0;
});
await page.waitForTimeout(200);
const cellBoxOf = async (grid) => await page.evaluate((g) => {
  const n = document.querySelector("#" + g + ' .gcells .gc[data-r="2"][data-c="0"]');
  const r = n.getBoundingClientRect();
  return { w: Math.round(r.width), h: Math.round(r.height) };
}, grid);

check("元データ・出力シートの両方に倍率がある",
  (await page.locator("#zoomSrc").count()) === 1 && (await page.locator("#zoomDst").count()) === 1);
check("はじめは100%", (await page.textContent("#zoomSrc .zl")) === "100%");
const z100 = await cellBoxOf("srcGrid");
await page.click('#zoomSrc button[data-z="1"]');
await page.waitForTimeout(200);
const z110 = await cellBoxOf("srcGrid");
check("＋で大きくなる", (await page.textContent("#zoomSrc .zl")) === "110%" && z110.w > z100.w,
  `${z100.w} → ${z110.w}`);
await page.click('#zoomSrc button[data-z="1"]');
await page.waitForTimeout(200);
check("さらに＋で125%", (await page.textContent("#zoomSrc .zl")) === "125%");

// 拡大したままでも、狙ったセルを正しく掴める（当たり判定が倍率でずれない）
await page.click('#srcGrid .gcells .gc[data-r="4"][data-c="2"]');
await page.waitForTimeout(150);
const zSel = await page.evaluate(() => window.__app.S.sel);
check("拡大中でも正しいセルを選べる",
  zSel.r1 === 4 && zSel.c1 === 2 && zSel.r2 === 4 && zSel.c2 === 2, JSON.stringify(zSel));
// ドラッグでの範囲選択も
const zb1 = await page.locator('#srcGrid .gcells .gc[data-r="1"][data-c="0"]').first().boundingBox();
const zb2 = await page.locator('#srcGrid .gcells .gc[data-r="4"][data-c="2"]').first().boundingBox();
await page.mouse.move(zb1.x + 10, zb1.y + 6);
await page.mouse.down();
await page.mouse.move(zb2.x + 20, zb2.y + 8, { steps: 8 });
await page.mouse.up();
const zDrag = await page.evaluate(() => window.__app.S.sel);
check("拡大中でもドラッグで範囲を選べる",
  zDrag.r1 === 1 && zDrag.c1 === 0 && zDrag.r2 === 4 && zDrag.c2 === 2, JSON.stringify(zDrag));

// 縮小と、下限で押せなくなること
await page.evaluate(() => window.__app.setZoom("src", 50, true));
await page.waitForTimeout(200);
const z50 = await cellBoxOf("srcGrid");
check("縮小すると小さくなる", z50.w < z100.w, `${z100.w} → ${z50.w}`);
check("下限では − が押せない", await page.isDisabled('#zoomSrc button[data-z="-1"]'));
await page.click('#srcGrid .gcells .gc[data-r="6"][data-c="3"]');
await page.waitForTimeout(150);
const zSmall = await page.evaluate(() => window.__app.S.sel);
check("縮小中でも正しいセルを選べる", zSmall.r1 === 6 && zSmall.c1 === 3, JSON.stringify(zSmall));
await page.evaluate(() => window.__app.setZoom("src", 200, true));
await page.waitForTimeout(200);
check("上限では ＋ が押せない", await page.isDisabled('#zoomSrc button[data-z="1"]'));

// 数字をクリックすると 100% に戻る
await page.click("#zoomSrc .zl");
await page.waitForTimeout(200);
check("数字をクリックで100%に戻る", (await page.textContent("#zoomSrc .zl")) === "100%"
  && (await cellBoxOf("srcGrid")).w === z100.w);

// 出力側も別々に効く
const dz100 = await cellBoxOf("dstGrid");
await page.click('#zoomDst button[data-z="-1"]');
await page.waitForTimeout(200);
const dz90 = await cellBoxOf("dstGrid");
check("出力シートの倍率は別に効く",
  (await page.textContent("#zoomDst .zl")) === "90%" && dz90.w < dz100.w
  && (await page.textContent("#zoomSrc .zl")) === "100%", `${dz100.w} → ${dz90.w}`);
check("倍率はブラウザに覚えておく（値だけ）",
  /"zoomDst":90/.test(await page.evaluate(() => localStorage.getItem("excel-extract-prefs-v1"))),
  await page.evaluate(() => localStorage.getItem("excel-extract-prefs-v1")));
await page.click("#zoomDst .zl");
await page.waitForTimeout(200);

// 出力シートの全画面表示
await page.click("#btnMaxDst");
await page.waitForTimeout(250);
check("出力シートも全画面にできる",
  !(await page.isVisible("#srcGrid")) && (await page.isVisible("#dstGrid")));
check("全画面の出力にもセルが描かれる",
  (await page.$$eval("#dstGrid .gc", (ns) => ns.length)) > 20,
  String(await page.$$eval("#dstGrid .gc", (ns) => ns.length)));
check("ボタンが戻す表示になる", (await page.textContent("#btnMaxDst")) === "⤡ 戻す");
await page.keyboard.press("Escape");
await page.waitForTimeout(250);
check("Escで戻る", await page.isVisible("#srcGrid"));

// 全画面は、どの並べ方から入っても画面いっぱいになること。
// 並べ方の指定に負けていたころは、ブック構成を畳んだ状態から全画面にすると
// 34px の列に、上下に並べた状態からだと元の高さのままの行に押し込まれていた
const maxedSize = async (which) => {
  await page.click(which === "src" ? "#btnMaxSrc" : "#btnMaxDst");
  await page.waitForTimeout(300);
  const r = await page.evaluate((w) => {
    const n = document.querySelector(w === "src" ? ".pane-src" : ".pane-dst");
    const b = n.getBoundingClientRect();
    const g = document.getElementById(w === "src" ? "srcGrid" : "dstGrid").getBoundingClientRect();
    return { pw: Math.round(b.width), ph: Math.round(b.height), gw: Math.round(g.width) };
  }, which);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  return r;
};
for (const [name, rail, stacked] of [
  ["そのまま", false, false], ["ブック構成をたたむ", true, false],
  ["上下に並べる", false, true], ["両方", true, true],
]) {
  await page.evaluate(([r, k]) => {
    document.body.classList.toggle("railoff", r);
    document.body.classList.toggle("stacked", k);
  }, [rail, stacked]);
  await page.waitForTimeout(150);
  const a = await maxedSize("src"), b = await maxedSize("dst");
  check(`${name}から全画面にしても画面いっぱいになる`,
    a.pw === 1440 && a.ph > 800 && a.gw > 1400 && b.pw === 1440 && b.ph > 800,
    JSON.stringify({ src: a, dst: b }));
}
await page.evaluate(() => { document.body.classList.remove("railoff", "stacked"); });
await page.waitForTimeout(200);

// 左右／上下の並べ替え
const sideBySide = await page.evaluate(() => {
  const s = document.querySelector(".pane-src").getBoundingClientRect();
  const d = document.querySelector(".pane-dst").getBoundingClientRect();
  return { sameRow: Math.abs(s.top - d.top) < 4, dRight: d.left > s.left };
});
check("はじめは左右に並ぶ", sideBySide.sameRow && sideBySide.dRight, JSON.stringify(sideBySide));
await page.click("#btnLayout");
await page.waitForTimeout(250);
const stacked = await page.evaluate(() => {
  const s = document.querySelector(".pane-src").getBoundingClientRect();
  const d = document.querySelector(".pane-dst").getBoundingClientRect();
  return { dBelow: d.top > s.top + 50, sameLeft: Math.abs(s.left - d.left) < 4, label: document.getElementById("btnLayout").textContent };
});
check("上下に並べ替えられる", stacked.dBelow && stacked.sameLeft, JSON.stringify(stacked));
check("ボタンの文言が戻す側になる", /左右に並べる/.test(stacked.label), stacked.label);
const stackedCells = await page.evaluate(() => {
  const w = document.getElementById("dstGrid"), G = window.__app.dstGrid;
  return {
    gc: w.querySelectorAll(".gc").length, h: w.clientHeight, w: w.clientWidth,
    win: G.win, rows: G.rows, cols: G.cols,
    inCells: w.querySelector(".gcells").children.length,
  };
});
check("並べ替えてもセルは描かれている", stackedCells.gc > 10, JSON.stringify(stackedCells));
await page.click("#btnLayout");
await page.waitForTimeout(250);
check("左右に戻せる",
  await page.evaluate(() => {
    const s = document.querySelector(".pane-src").getBoundingClientRect();
    const d = document.querySelector(".pane-dst").getBoundingClientRect();
    return Math.abs(s.top - d.top) < 4 && d.left > s.left;
  }));

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
// 置かれるのは手順書と画面の好みだけ（どちらもセルの値は含まない）
check("localStorage に置くのは画面設定だけ",
  Object.keys(storage.ls).every((k) => k === "excel-extract-prefs-v1"),
  Object.keys(storage.ls).join(","));
check("画面設定には並べ方と折りたたみしか入らない",
  Object.keys(JSON.parse(storage.ls["excel-extract-prefs-v1"] || "{}"))
    .every((k) => ["stacked", "foldSheets", "foldBlocks", "rail", "splitW", "splitH",
      "zoomSrc", "zoomDst"].indexOf(k) >= 0),
  storage.ls["excel-extract-prefs-v1"]);
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
  await page.evaluate(() => window.__app.S.files.length === 0 && window.__app.S.sheets.length === 0));
check("手順書はブラウザに残さない",
  (await page.evaluate(() => localStorage.getItem("excel-extract-recipes-v1"))) === null);

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
await p2.click("#btnSample2");
await p2.waitForSelector("#srcGrid:not([hidden])");
check("Artifact 断片版も起動する",
  (await p2.$$eval("#sheetList .sheet-item", (ns) => ns.length)) === 3 && errors2.length === 0,
  errors2.slice(0, 2).join(" | "));

await browser.close();
server.close();

console.log(`\n${checks.filter((c) => c.ok).length}/${checks.length} 件成功`);
if (fails.length) { console.error("\n失敗:\n- " + fails.join("\n- ")); process.exit(1); }
console.log("すべて成功");
