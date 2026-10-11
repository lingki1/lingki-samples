// 第四版起的「重复发票：对比、选保留哪张、另一张怎么处理」功能测试：三组重复各用一种处理方式，并看原件有没有被动。
// 用法：node dup_test.mjs <v4 或更新的那一版.html>
// 截图和导出的表格存到「输出」文件夹。依赖和装法见 _common.mjs 开头。全程掐断往外的网络。
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { htmlArg, launch, outDir, sampleNames, serve } from './_common.mjs';
const site = await serve(htmlArg('node dup_test.mjs <v4 或更新的那一版.html>'));
const names = sampleNames(); const OUT = outDir();
const browser = await launch(); const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
await ctx.route((u) => !site.isLocal(u.href) && /^https?:/.test(u.href), (r) => r.abort());
const page = await ctx.newPage(); const errs = []; page.on('pageerror', (e) => errs.push(e.message)); await page.goto(`${site.origin}/app.html`);
await page.evaluate(async (names) => { localStorage.clear(); const root = await navigator.storage.getDirectory(); try { await root.removeEntry('发票', { recursive: true }); } catch {}
  const d = await root.getDirectoryHandle('发票', { create: true });
  for (const n of names) { const b = await (await fetch('/发票/' + encodeURIComponent(n))).blob(); const fh = await d.getFileHandle(n, { create: true }); const w = await fh.createWritable(); await w.write(b); await w.close(); }
  await loadFolder(d); }, names);
const groups = await page.evaluate(() => [...new Set(items.filter((i) => i.group).map((i) => i.group))].map((g) => g.map((x) => ({ id: x.id, name: x.name, status: x.status, amount: x.amount }))));
console.log('三组重复（处理之前）：'); for (const g of groups) console.log('  ', g.map((x) => `${x.name}[${x.status}]`).join('  ↔  '));
const choose = async (g, act, switchKeeper) => { const dup = g.find((x) => x.status === 'dup');
  await page.click('.tab[data-filter="dup"], [data-filter="dup"]').catch(() => {}); await page.evaluate((id) => openCompare(items[id]), dup.id); await page.waitForSelector('#cmp[open]');
  if (switchKeeper) await page.click(`.cmp-pick[data-id="${dup.id}"]`); await page.check(`[name=cmpAct][value="${act}"]`); return dup; };
// 第 1 组：改成保留原来被判成重复的那张，另一张「不放进整理结果」
const d1 = await choose(groups[0], 'skip', true); await page.screenshot({ path: join(OUT, 'dup-compare.png') }); await page.click('#cmpOk');
// 第 2 组：「它们不是重复，都当正常发票保留」
await choose(groups[1], 'both', false); await page.click('#cmpOk');
// 第 3 组：不动（默认：另一张放进「重复发票」）
await page.waitForTimeout(200); await page.screenshot({ path: join(OUT, 'dup-tab.png') });
const after = await page.evaluate(() => ({ st: items.filter((i) => i.group).map((i) => `${i.name}[${i.status}${i.skip ? ',不放进来' : ''}]`), ok: items.filter((i) => i.status === 'ok').length, total: items.filter((i) => i.status === 'ok').reduce((s, i) => s + Math.round(i.amount * 100), 0) / 100 }));
console.log('处理之后：', after.st.join('  ')); console.log(`正常发票 ${after.ok} 张，合计 ${after.total.toFixed(2)}；应为 81 张、${(29914.93 + groups[1][0].amount).toFixed(2)}`);
await page.click('#run'); await page.waitForSelector('#done[open]', { timeout: 60000 }); await page.screenshot({ path: join(OUT, 'dup-done.png') });
const res = await page.evaluate(async () => { const root = await navigator.storage.getDirectory(); const d = await root.getDirectoryHandle('发票'); const tree = {}; let out, src = 0;
  async function walk(dir, p) { for await (const [n, h] of dir.entries()) { if (h.kind === 'directory') await walk(h, p + n + '/'); else (tree[p] = tree[p] || []).push(n); } }
  for await (const [n, h] of d.entries()) { if (h.kind === 'directory') { out = h; await walk(h, ''); } else src++; }
  const x = await (await out.getFileHandle('发票汇总表.xlsx')).getFile(); const u = new Uint8Array(await x.arrayBuffer()); let s = ''; for (let i = 0; i < u.length; i += 8192) s += String.fromCharCode(...u.subarray(i, i + 8192));
  return { src, tree, xlsx: btoa(s), done: document.querySelector('#doneList').innerText }; });
writeFileSync(join(OUT, 'dup-test.xlsx'), Buffer.from(res.xlsx, 'base64'));
console.log('原文件夹里的原件还有', res.src, '个（应为 85）'); console.log('完成窗口写的：', res.done.replace(/\n/g, ' ｜ '));
console.log('输出各文件夹：', Object.entries(res.tree).map(([k, v]) => `${k || '(根)'} ${v.length}`).join('，')); console.log('「重复发票」里：', res.tree['重复发票/'] || '（没有这个文件夹）');
console.log('第 1 组被选为「不放进来」的那张', groups[0].find((x) => x.id !== d1.id).name, '在输出里出现了吗：', Object.values(res.tree).flat().length, '个输出文件（应为 85 − 1 + 表格 1 = 85）');
console.log('页面报错：', errs.length ? errs : '无'); await browser.close(); site.close();
