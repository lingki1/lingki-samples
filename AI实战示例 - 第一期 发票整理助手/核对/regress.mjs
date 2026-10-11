// 回归检查：把某一版「发票整理助手.html」在浏览器内核里跑一遍，85 个样例的每一格和标准答案比；同时记下它有没有往外联网。
// 用法：node regress.mjs <某一版.html> [--offline] [--shot=前缀]
//   --offline   把所有往外的网络请求都掐掉（第一版要联网取组件，掐掉就读不了；第二版起应当照常全对）
//   --shot=前缀  顺手截三张图，存到「输出」文件夹
// 依赖和装法见 _common.mjs 开头。
import { join } from 'node:path';
import { htmlArg, launch, outDir, readTruth, sampleNames, serve } from './_common.mjs';
const site = await serve(htmlArg('node regress.mjs <某一版.html> [--offline] [--shot=前缀]'));
const T = readTruth(); const truth = Object.fromEntries(T.items.map((t) => [t.file, t]));
const names = sampleNames();
const OFF = process.argv.includes('--offline'); const SHOT = (process.argv.find((a) => a.startsWith('--shot=')) || '').slice(7);
const browser = await launch(); const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
const outside = []; ctx.on('request', (r) => { if (!site.isLocal(r.url()) && !/^(blob:|data:|about:)/.test(r.url())) outside.push(r.url().slice(0, 90)); });
if (OFF) await ctx.route((u) => !site.isLocal(u.href) && /^https?:/.test(u.href), (r) => r.abort());
const page = await ctx.newPage(); const errs = []; page.on('pageerror', (e) => errs.push(e.message)); await page.goto(`${site.origin}/app.html`);
if (SHOT) await page.screenshot({ path: join(outDir(), `${SHOT}-empty.png`) });
const got = await page.evaluate(async (names) => {
  const root = await navigator.storage.getDirectory(); try { await root.removeEntry('发票', { recursive: true }); } catch {}
  const d = await root.getDirectoryHandle('发票', { create: true });
  for (const n of names) { const b = await (await fetch('/发票/' + encodeURIComponent(n))).blob(); const fh = await d.getFileHandle(n, { create: true }); const w = await fh.createWritable(); await w.write(b); await w.close(); }
  await loadFolder(d);
  return { msg: document.querySelector('#msg') && !document.querySelector('#msg').hidden ? document.querySelector('#msg').textContent : '', items: items.map((i) => ({ name: i.name, status: i.status, date: i.date, seller: i.seller, buyer: i.buyer, amount: i.amount, tax: i.tax, net: i.net, number: i.number, item: i.item, folder: i.folder, newName: i.newName, dupOf: i.dupOf && i.dupOf.name })) };
}, names);
const bad = []; const st = {}; let total = 0;
if (got.items.length !== T.items.length) bad.push(`只读到 ${got.items.length} 个文件，标准答案里有 ${T.items.length} 个`);
for (const g of got.items) { const t = truth[g.name]; st[g.status] = (st[g.status] || 0) + 1; if (!t) { bad.push(`不认识的文件 ${g.name}`); continue; }
  if (t.kind === 'not-invoice') { if (g.status !== 'other') bad.push(`行程单状态 ${g.status}`); continue; }
  if (t.kind === 'image-only') { if (g.status !== 'check' || g.amount != null) bad.push(`纯图片那张 ${g.status} ${g.amount}`); continue; }
  for (const [k, v] of [['date', t.date], ['seller', t.seller], ['buyer', t.buyer], ['number', t.number], ['amount', t.total], ['tax', t.tax]]) if (String(g[k]) !== String(v)) bad.push(`${g.name} ${k}: ${g[k]} ≠ ${v}`);
  if (g.status === 'ok') { total += Math.round(g.amount * 100); if (!g.newName.startsWith(`${t.date}_${t.seller}_${t.total.toFixed(2)}元`)) bad.push(`${g.name} 新名字 ${g.newName}`); } else if (g.status !== 'dup') bad.push(`${g.name} 状态 ${g.status}`); }
console.log(`${OFF ? '断网' : '联网'}：读到 ${got.items.length} 个`, JSON.stringify(st), `正常发票合计 ${(total / 100).toFixed(2)}（标准答案 ${T.total.toFixed(2)}）`);
console.log('页面提示：', got.msg || '（无）', '| 页面报错：', errs.length ? errs : '无', '| 往外联网的请求：', outside.length ? outside : '没有');
console.log(bad.length ? `不对的 ${bad.length} 处：\n  ` + bad.slice(0, 20).join('\n  ') : '逐格比对：全对');
if (SHOT) { await page.waitForTimeout(400); await page.screenshot({ path: join(outDir(), `${SHOT}-loaded.png`) }); await page.screenshot({ path: join(outDir(), `${SHOT}-full.png`), fullPage: true }); }
await browser.close(); site.close();
