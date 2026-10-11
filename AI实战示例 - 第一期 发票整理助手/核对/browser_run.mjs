// 在浏览器内核里把第一版「发票整理助手.html」跑三遍：正常联网、混进 OFD 和图片、连不上放组件的那个网站。
// 第一版读 PDF 的组件是每次从 cdnjs.cloudflare.com 取的，所以第 1、2 遍要能上网；第二版起组件内置了，第 3 遍对它们没有意义。
// 用法：node browser_run.mjs <v1 那一版.html>
// 截图存到「输出」文件夹。依赖和装法见 _common.mjs 开头。
import { join } from 'node:path';
import { htmlArg, launch, outDir, sampleNames, serve } from './_common.mjs';
const site = await serve(htmlArg('node browser_run.mjs <v1 那一版.html>'));
const names = sampleNames(); const OUT = outDir();
const URL0 = `${site.origin}/app.html`;
const browser = await launch();
const fill = async (page, extra = []) => page.evaluate(async ({ names, extra }) => {
  const root = await navigator.storage.getDirectory(); try { await root.removeEntry('发票', { recursive: true }); } catch {}
  const d = await root.getDirectoryHandle('发票', { create: true });
  for (const n of names) { const b = await (await fetch('/发票/' + encodeURIComponent(n))).blob(); await writeFile(d, n, b); }
  for (const [n, text] of extra) await writeFile(d, n, new Blob([text]));
  const t0 = performance.now(); await loadFolder(d); const ms = Math.round(performance.now() - t0);
  const count = (s) => items.filter((i) => i.status === s).length;
  return { ms, n: items.length, ok: count('ok'), dup: count('dup'), check: count('check'), other: count('other'), msg: document.querySelector('#msg').hidden ? '' : document.querySelector('#msg').textContent,
    total: items.filter((i) => i.status === 'ok').reduce((s, i) => s + i.amount, 0).toFixed(2), extras: items.filter((i) => !/\.pdf$/i.test(i.name)).map((i) => [i.name, i.status, i.folder]) };
}, { names, extra });

// 1. 正常联网
{ const page = await browser.newPage({ viewport: { width: 1440, height: 900 } }); page.on('pageerror', (e) => console.log('PAGE ERROR', e.message)); await page.goto(URL0);
  await page.screenshot({ path: join(OUT, 'v1-empty.png') });
  console.log('联网：', JSON.stringify(await fill(page))); await page.waitForTimeout(300); await page.screenshot({ path: join(OUT, 'v1-loaded.png') });
  await page.evaluate(() => document.querySelector('#rows .fname').click()); await page.waitForTimeout(800); await page.screenshot({ path: join(OUT, 'v1-preview.png') }); await page.close(); }
// 2. 混进一个 OFD 和一张 JPG（内容是假的，只看它把这两种文件放到哪、怎么说；两个假文件的内容故意不一样，免得被它当成同一个文件的两份）
{ const page = await browser.newPage(); await page.goto(URL0); console.log('混进 OFD 和 JPG：', JSON.stringify(await fill(page, [['数电票.ofd', '假的 OFD'], ['拍的发票.jpg', '假的 JPG'], ['26112000000999999999.xml', '<x/>']]))); await page.close(); }
// 3. 连不上放组件的那个网站
{ const ctx = await browser.newContext(); await ctx.route(/cdnjs\.cloudflare\.com/, (r) => r.abort()); const page = await ctx.newPage(); await page.goto(URL0); console.log('连不上组件网站：', JSON.stringify(await fill(page))); await page.screenshot({ path: join(OUT, 'v1-offline.png') }); await ctx.close(); }
await browser.close(); site.close();
