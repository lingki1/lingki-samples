// 把某一版「发票整理助手.html」放进模拟的手机屏里截两张图，并量一下页面比屏幕宽出多少。
// 不是真手机：只是手机的屏幕宽度（412×915）、触屏和安卓 Chrome 的浏览器标识。
// 用法：node phone_shot.mjs <某一版.html> [截图文件名的前缀]
// 截图存到「输出」文件夹。依赖和装法见 _common.mjs 开头。
import { join } from 'node:path';
import { htmlArg, launch, outDir, sampleNames, serve } from './_common.mjs';
const site = await serve(htmlArg('node phone_shot.mjs <某一版.html> [截图文件名的前缀]'));
const names = sampleNames(); const OUT = outDir(); const pre = process.argv.slice(2).filter((a) => !a.startsWith('--'))[1] || 'phone';
const URL0 = `${site.origin}/app.html`;
const browser = await launch();
const phone = { viewport: { width: 412, height: 915 }, deviceScaleFactor: 2.1626, isMobile: true, hasTouch: true, userAgent: 'Mozilla/5.0 (Linux; Android 15; Pixel 9) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Mobile Safari/537.36' };
// 1. 手机上打开时看到的第一屏（手机浏览器大多没有「选文件夹」这个功能，所以这里把它拿掉）
{ const ctx = await browser.newContext(phone); await ctx.addInitScript(() => { delete window.showDirectoryPicker; }); const page = await ctx.newPage(); await page.goto(URL0); await page.waitForTimeout(400);
  await page.screenshot({ path: join(OUT, `${pre}-start.png`) }); console.log('首屏：页面宽度', await page.evaluate(() => document.documentElement.scrollWidth), '屏幕宽度', await page.evaluate(() => window.innerWidth)); await ctx.close(); }
// 2. 假如这台手机能选文件夹，读完 85 张之后的样子
{ const ctx = await browser.newContext(phone); const page = await ctx.newPage(); await page.goto(URL0);
  const ok = await page.evaluate(async (names) => { try { const root = await navigator.storage.getDirectory(); try { await root.removeEntry('发票', { recursive: true }); } catch {}
    const d = await root.getDirectoryHandle('发票', { create: true });
    for (const n of names) { const b = await (await fetch('/发票/' + encodeURIComponent(n))).blob(); const fh = await d.getFileHandle(n, { create: true }); const w = await fh.createWritable(); await w.write(b); await w.close(); }
    await loadFolder(d); return items.length; } catch (e) { return String(e); } }, names);
  await page.waitForTimeout(500); await page.screenshot({ path: join(OUT, `${pre}-loaded.png`) }); console.log('读完之后：', ok, '页面宽度', await page.evaluate(() => document.documentElement.scrollWidth)); await ctx.close(); }
await browser.close(); site.close();
