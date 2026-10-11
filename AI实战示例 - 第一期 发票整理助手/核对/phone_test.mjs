// 第五版在模拟的手机屏里走一遍新路：选文件 → 读 85 张 → 看原件 → 整理并打包下载。
// 不是真手机：只是手机的屏幕宽度（412）、触屏、安卓 Chrome 的浏览器标识，并拿掉了「选文件夹」这个功能。
// 用法：node phone_test.mjs <v5 那一版.html>
// 截图和下载到的压缩包存到「输出」文件夹。依赖和装法见 _common.mjs 开头。全程掐断往外的网络。
import { join } from 'node:path';
import { SAMPLES, htmlArg, launch, outDir, readTruth, sampleNames, serve } from './_common.mjs';
const site = await serve(htmlArg('node phone_test.mjs <v5 那一版.html>'));
const T = readTruth(); const truth = Object.fromEntries(T.items.map((t) => [t.file, t]));
const files = sampleNames().map((n) => join(SAMPLES, n)); const OUT = outDir();
const browser = await launch();
const ctx = await browser.newContext({ viewport: { width: 412, height: 915 }, deviceScaleFactor: 2.1626, isMobile: true, hasTouch: true, acceptDownloads: true, userAgent: 'Mozilla/5.0 (Linux; Android 15; Pixel 9) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Mobile Safari/537.36' });
await ctx.addInitScript(() => { delete window.showDirectoryPicker; });
const outside = []; ctx.on('request', (r) => { if (!site.isLocal(r.url()) && !/^(blob:|data:|about:)/.test(r.url())) outside.push(r.url().slice(0, 80)); });
await ctx.route((u) => !site.isLocal(u.href) && /^https?:/.test(u.href), (r) => r.abort());
const page = await ctx.newPage(); const errs = []; page.on('pageerror', (e) => errs.push(e.message)); await page.goto(`${site.origin}/app.html`); await page.waitForTimeout(300);
const wide = () => page.evaluate(() => `页面宽 ${document.documentElement.scrollWidth} / 屏幕宽 ${window.innerWidth}`);
await page.screenshot({ path: join(OUT, 'phone-start.png') }); console.log('首屏：', await wide(), '| 主按钮上写的：', await page.evaluate(() => [...document.querySelectorAll('#start button, #start a')].filter((b) => b.offsetParent).map((b) => b.innerText.trim()).join(' / ')));
await page.setInputFiles('#fileInput', files); await page.waitForFunction(() => typeof items !== 'undefined' && items.length === 85 && !document.querySelector('#work').hidden, null, { timeout: 120000 });
await page.waitForTimeout(500); await page.screenshot({ path: join(OUT, 'phone-loaded.png') }); console.log('读完：', await wide());
const got = await page.evaluate(() => items.map((i) => ({ name: i.name, status: i.status, date: i.date, seller: i.seller, amount: i.amount, tax: i.tax, number: i.number, newName: i.newName })));
const bad = []; const st = {}; let total = 0;
for (const g of got) { const t = truth[g.name]; st[g.status] = (st[g.status] || 0) + 1; if (!t) { bad.push('不认识 ' + g.name); continue; } if (t.kind === 'not-invoice' || t.kind === 'image-only') continue;
  for (const [k, v] of [['date', t.date], ['seller', t.seller], ['number', t.number], ['amount', t.total], ['tax', t.tax]]) if (String(g[k]) !== String(v)) bad.push(`${g.name} ${k}: ${g[k]} ≠ ${v}`); if (g.status === 'ok') total += Math.round(g.amount * 100); }
console.log('选文件这条路：读到', got.length, JSON.stringify(st), '合计', (total / 100).toFixed(2), bad.length ? `不对 ${bad.length} 处 ${bad.slice(0, 5)}` : '逐格全对');
// 看原件（它说手机上由软件自己把发票画出来）
const btn = page.locator('#rows [data-act="preview"]').first(); await btn.scrollIntoViewIfNeeded(); await btn.click(); await page.waitForTimeout(1500); await page.screenshot({ path: join(OUT, 'phone-preview.png') });
console.log('原件预览：', await page.evaluate(() => { const b = document.querySelector('#previewBody'); if (!b) return '找不到预览区'; const c = b.querySelectorAll('canvas'), i = b.querySelectorAll('iframe, img'); return `画布 ${c.length} 个${c.length ? `（${c[0].width}×${c[0].height}）` : ''}，内嵌框或图 ${i.length} 个`; }));
await page.locator('#previewClose').click().catch(() => {});
const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), page.locator('#run').click()]); await dl.saveAs(join(OUT, 'phone-out.zip')); await page.waitForTimeout(500); await page.screenshot({ path: join(OUT, 'phone-done.png') });
console.log('打包下载：', dl.suggestedFilename(), '（已另存为 输出/phone-out.zip，解开后可以用 check_run.py 核）| 完成窗口：', await page.evaluate(() => document.querySelector('#doneLead').textContent));
console.log('页面报错：', errs.length ? errs : '无', '| 往外联网：', outside.length ? outside : '没有'); await browser.close(); site.close();
