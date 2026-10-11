// 生成样例的脚本：造出这个文件夹里那 85 个「样例发票」PDF 和「标准答案.json」。
//
// 为什么要造：拿不到可以公开的真发票，所以视频里用的发票全部是自己造的。公司名、号码、金额都是编的
// （如果和哪家真实的店或公司重名，纯属巧合），每页压着「样例 · 非真实发票」的水印，不放章、不放二维码、不照官方版式排。
// 字段的叫法和真发票一样（发票号码、开票日期、购买方 / 销售方名称、价税合计），这样读发票的工具面对的是同一类问题。
// 里面故意埋了三种坑：3 张重复的（同一张又存了一份）、1 张整页是图片读不出字的、1 份根本不是发票的行程单。
//
// 依赖：
//   1. Node.js 18 或更新的版本
//   2. playwright（用它带的浏览器把网页存成 PDF）。在这个文件夹里运行：
//        npm install playwright
//        npx playwright install chromium
//      已经在别处装过 playwright 的，可以用环境变量 PLAYWRIGHT_PATH 指过去；
//      想用电脑上现成的 Chrome / Edge，用环境变量 CHROME_PATH 指向它的 exe。
//   3. 字体用的是 Windows 自带的宋体和微软雅黑。别的系统没有这两种字体，造出来的版面会略有差别。
//
// 怎么跑：
//        node 生成样例的脚本.mjs            → 造到旁边新建的「重新生成」文件夹里
//        node 生成样例的脚本.mjs 某个文件夹   → 造到你指定的文件夹里（必须是空的或还不存在，脚本不会删任何东西）
//
// 随机数的种子是写死的，所以每次造出来的发票内容、文件名、标准答案都一样。
// 但 PDF 文件本身和仓库里那 85 个不会逐字节相同：PDF 里记着生成的时间，浏览器版本不同排出来也可能差一点。
import { createRequire } from 'node:module';
import { copyFileSync, existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const HERE = dirname(fileURLToPath(import.meta.url));
function loadChromium() {
  for (const name of [process.env.PLAYWRIGHT_PATH, 'playwright', 'playwright-core'].filter(Boolean)) { try { return require(name).chromium; } catch { /* 试下一个 */ } }
  console.error('没找到 playwright。先在这个文件夹里运行：npm install playwright，再运行：npx playwright install chromium'); process.exit(1);
}
const chromium = loadChromium();
const OUT = resolve(process.argv[2] || join(HERE, '重新生成'));
if (existsSync(OUT) && readdirSync(OUT).length) { console.error(`「${OUT}」里已经有东西了。换一个空文件夹，或者自己先清空它（这个脚本不删文件）。`); process.exit(1); }
mkdirSync(OUT, { recursive: true });

let seed = 20261010; const rnd = () => { seed |= 0; seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
const pick = (a) => a[Math.floor(rnd() * a.length)]; const int = (a, b) => a + Math.floor(rnd() * (b - a + 1)); const pad = (n, w = 2) => String(n).padStart(w, '0');
const digits = (n) => Array.from({ length: n }, () => int(0, 9)).join('');

const BUYER = { name: '青岚设计工作室', id: '92SAMPLE0000000001' };
// [销售方, 假税号, 项目, 税率, 金额下限, 金额上限, 文件名里的简称]
const SELLERS = [
  ['青柚出行科技有限公司', '91SAMPLE0000000011', '*运输服务*客运服务费', 0.03, 9, 96, '青柚出行'],
  ['海棠路小馆餐饮店', '92SAMPLE0000000012', '*餐饮服务*餐费', 0.01, 28, 460, '海棠路小馆'],
  ['云栖客栈酒店管理有限公司', '91SAMPLE0000000013', '*住宿服务*住宿费', 0.06, 268, 1380, '云栖客栈'],
  ['方格纸文具商行', '92SAMPLE0000000014', '*纸制品*复印纸', 0.13, 35, 620, '方格纸文具'],
  ['南风通信服务有限公司', '91SAMPLE0000000015', '*电信服务*通讯服务费', 0.09, 59, 199, '南风通信'],
  ['向阳加油站', '92SAMPLE0000000016', '*汽油*92号车用汽油', 0.13, 150, 420, '向阳加油站'],
  ['飞燕速运有限公司', '91SAMPLE0000000017', '*物流辅助服务*快递费', 0.06, 8, 86, '飞燕速运'],
  ['栗子云计算有限公司', '91SAMPLE0000000018', '*信息技术服务*云服务费', 0.06, 99, 2400, '栗子云'],
  ['梧桐里停车管理有限公司', '91SAMPLE0000000019', '*经营租赁*停车费', 0.09, 5, 60, '梧桐里停车'],
  ['拾光图文快印店', '92SAMPLE0000000020', '*印刷品*打印制作费', 0.01, 20, 880, '拾光图文'],
  ['半山咖啡馆', '92SAMPLE0000000021', '*餐饮服务*餐饮服务', 0.01, 18, 150, '半山咖啡'],
  ['北辰办公家具有限公司', '91SAMPLE0000000022', '*家具*办公椅', 0.13, 399, 3600, '北辰家具'],
];
const UP = (n) => { const d = '零壹贰叁肆伍陆柒捌玖', u = ['', '拾', '佰', '仟'], g = ['', '万']; let [i, f] = n.toFixed(2).split('.'); let s = ''; i = String(parseInt(i, 10));
  if (i === '0') s = '零'; else { const gs = []; while (i.length) { gs.unshift(i.slice(-4)); i = i.slice(0, -4); }
    gs.forEach((gr, gi) => { let part = '', zero = false; for (let k = 0; k < gr.length; k++) { const c = +gr[k]; if (!c) { zero = true; continue; } if (zero && (part || s)) part += '零'; zero = false; part += d[c] + u[gr.length - 1 - k]; } if (part) s += part + g[gs.length - 1 - gi]; }); }
  s += '圆'; const j = +f[0], fen = +f[1]; if (!j && !fen) return `${s}整`; if (j) s += `${d[j]}角`; else s += '零'; return fen ? `${s}${d[fen]}分` : `${s}整`; };
const money = (n, style) => (style === 0 ? `¥${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : style === 1 ? `￥${n.toFixed(2)}` : `¥ ${n.toFixed(2)}`);
const dateText = (y, m, d, style) => (style === 0 ? `${y}年${pad(m)}月${pad(d)}日` : style === 1 ? `${y}-${pad(m)}-${pad(d)}` : `${y}年${m}月${d}日`);

const CSS = `@page { size: 210mm 140mm; margin: 0 } html, body { margin: 0 } body { width: 210mm; height: 140mm; position: relative; overflow: hidden; font-family: "SimSun", "Microsoft YaHei", serif; color: #3a2a20; font-size: 10.5pt; }
  .wm { position: absolute; left: 0; right: 0; top: 46mm; text-align: center; font: 700 38pt "Microsoft YaHei", sans-serif; color: rgba(200, 30, 60, .16); transform: rotate(-16deg); letter-spacing: 4pt; white-space: nowrap; }
  .pg { position: absolute; inset: 9mm 11mm; } h1 { margin: 0; text-align: center; font: 700 19pt "Microsoft YaHei", sans-serif; letter-spacing: 6pt; color: #7a3b1e; } h1 small { display: block; font-size: 8.5pt; letter-spacing: 1pt; font-weight: 400; margin-top: 1mm; }
  .meta { position: absolute; right: 0; top: 0; text-align: left; line-height: 1.7; font-size: 10pt; } .box { border: .35mm solid #9b5a3a; }
  .two { display: flex; margin-top: 7mm; } .two > div { flex: 1; padding: 2.5mm 3mm; line-height: 1.75; } .two > div + div { border-left: .35mm solid #9b5a3a; } .two b { color: #7a3b1e; }
  table { border-collapse: collapse; width: 100%; } td, th { padding: 1.6mm 2mm; text-align: left; font-weight: 400; } th { color: #7a3b1e; border-bottom: .2mm solid #9b5a3a; } .r { text-align: right; }
  .items { border-top: 0; min-height: 30mm; } .sum { border-top: 0; padding: 2.5mm 3mm; display: flex; gap: 8mm; } .sum b { color: #7a3b1e; font-weight: 400; } .foot { margin-top: 3mm; display: flex; gap: 14mm; font-size: 9.5pt; }`;

// 三种排法（tpl 0 / 1 / 2）：字段的位置、冒号、金额和日期的写法各不一样
function html(v) {
  const net = +(v.total / (1 + v.rate)).toFixed(2); const tax = +(v.total - net).toFixed(2); v.tax = tax;
  const c = v.tpl === 1 ? ':' : '：'; const M = (n) => money(n, v.tpl); const D = dateText(v.y, v.m, v.d, v.tpl);
  const meta = `<div class="meta">发票号码${c}${v.number}<br>开票日期${c}${D}</div>`;
  const row = `<tr><td>${v.item}</td><td class="r">1</td><td class="r">${net.toFixed(2)}</td><td class="r">${net.toFixed(2)}</td><td class="r">${Math.round(v.rate * 100)}%</td><td class="r">${tax.toFixed(2)}</td></tr>`;
  const head = '<tr><th>项目名称</th><th class="r">数量</th><th class="r">单价</th><th class="r">金额</th><th class="r">税率</th><th class="r">税额</th></tr>';
  let body;
  if (v.tpl === 0) body = `<h1>电子发票<small>（普通发票）</small></h1>${meta}<div class="two box"><div><b>购买方信息</b><br>名称${c}${v.buyer}<br>统一社会信用代码/纳税人识别号${c}${v.buyerId}</div><div><b>销售方信息</b><br>名称${c}${v.seller}<br>统一社会信用代码/纳税人识别号${c}${v.sellerId}</div></div>
    <div class="items box"><table>${head}${row}</table></div><div class="sum box"><span><b>价税合计（大写）</b> ${UP(v.total)}</span><span><b>（小写）</b> ${M(v.total)}</span></div><div class="foot"><span>备注${c}</span><span>开票人${c}${v.clerk}</span></div>`;
  else if (v.tpl === 1) body = `<h1>电子发票</h1>${meta}<div class="two box"><div>购买方名称${c} ${v.buyer}<br>纳税人识别号${c} ${v.buyerId}</div><div>销售方名称${c} ${v.seller}<br>纳税人识别号${c} ${v.sellerId}</div></div>
    <div class="items box"><table>${head}${row}<tr><td>合计</td><td></td><td></td><td class="r">${M(net)}</td><td></td><td class="r">${M(tax)}</td></tr></table></div><div class="sum box"><span>价税合计(小写)${c} ${M(v.total)}</span></div><div class="foot"><span>开票人${c} ${v.clerk}</span></div>`;
  else body = `<h1>电子发票<small>（普通发票）</small></h1><div class="box" style="margin-top:5mm"><table><tr><td>发票号码</td><td>${v.number}</td><td>开票日期</td><td>${D}</td></tr><tr><td>购买方</td><td colspan="3">${v.buyer}　${v.buyerId}</td></tr><tr><td>销售方</td><td colspan="3">${v.seller}　${v.sellerId}</td></tr></table></div>
    <div class="items box"><table>${head}${row}</table></div><div class="box" style="border-top:0"><table><tr><td>价税合计（大写）</td><td>${UP(v.total)}</td><td>价税合计（小写）</td><td class="r">${M(v.total)}</td></tr></table></div><div class="foot"><span>开票人　${v.clerk}</span></div>`;
  return `<!doctype html><meta charset="utf-8"><style>${CSS}</style><body><div class="pg">${body}</div><div class="wm">样例 · 非真实发票</div></body>`;
}

// 80 张不重样的发票，开票日期都在 2026 年 9 月；号码 20 位，和数电发票的号码一样长
const inv = []; const used = new Set();
for (let i = 0; i < 80; i++) { const s = pick(SELLERS); let number; do { number = `26${pick(['11', '31', '33', '44', '51'])}2000000${digits(9)}`; } while (used.has(number)); used.add(number);
  inv.push({ number, y: 2026, m: 9, d: int(1, 30), seller: s[0], sellerId: s[1], item: s[2], rate: s[3], total: +(int(s[4] * 100, s[5] * 100) / 100).toFixed(2), short: s[6], buyer: BUYER.name, buyerId: BUYER.id, tpl: pick([0, 0, 1, 1, 2]), clerk: pick(['周', '林', '许', '唐', '宋', '孟']) + pick(['一', '宁', '可', '然', '舟']), hh: int(8, 22), mi: int(0, 59), ss: int(0, 59) }); }

// 文件名学的是「下载下来就这样、看不出是什么」那几种写法
const names = new Set(); const uniq = (n) => { let k = n, i = 1; while (names.has(k.toLowerCase())) { k = n.replace(/\.pdf$/i, ` (${i++}).pdf`); } names.add(k.toLowerCase()); return k; };
const fileName = (v) => { const r = rnd(); if (r < 0.4) return `dzfp_${v.number}_2026${pad(v.m)}${pad(v.d)}${pad(v.hh)}${pad(v.mi)}${pad(v.ss)}.pdf`; if (r < 0.6) return `${v.number}.pdf`; if (r < 0.75) return '电子发票.pdf'; if (r < 0.85) return `invoice_${digits(8)}.pdf`; return `${v.short}-电子发票.pdf`; };

const browser = await chromium.launch(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}); const page = await browser.newPage();
const truth = [];
for (const v of inv) { await page.setContent(html(v)); v.file = uniq(fileName(v)); await page.pdf({ path: join(OUT, v.file), width: '210mm', height: '140mm', printBackground: true });
  truth.push({ file: v.file, kind: 'invoice', number: v.number, date: `${v.y}-${pad(v.m)}-${pad(v.d)}`, seller: v.seller, buyer: v.buyer, total: v.total, tax: v.tax, item: v.item, template: v.tpl }); }

// 埋的坑：三张重复的（同一张票又下了一遍）、一张只有图片没有文字的、一份根本不是发票的
for (const k of [7, 31, 58]) { const src = truth[k]; const f = uniq(pick(['电子发票.pdf', `${inv[k].short}-电子发票.pdf`, `${src.number}.pdf`])); copyFileSync(join(OUT, src.file), join(OUT, f)); truth.push({ ...src, file: f, kind: 'duplicate', duplicateOf: src.file }); }
{ const v = { ...inv[12], number: `26332000000${digits(9)}`, d: 18, total: 126.5 }; await page.setViewportSize({ width: 1240, height: 827 }); await page.setContent(html(v)); const png = await page.screenshot({ type: 'png' });
  await page.setContent(`<!doctype html><style>@page{size:210mm 140mm;margin:0}html,body{margin:0}img{width:210mm;height:140mm;display:block}</style><img src="data:image/png;base64,${png.toString('base64')}">`); const f = uniq('微信图片_20260918213045.pdf'); await page.pdf({ path: join(OUT, f), width: '210mm', height: '140mm', printBackground: true });
  truth.push({ file: f, kind: 'image-only', number: v.number, date: '2026-09-18', seller: v.seller, buyer: v.buyer, total: v.total, note: '整页是一张图，没有文字层；工具应该说读不了，不该猜' }); }
{ await page.setContent(`<!doctype html><meta charset="utf-8"><style>@page{size:A4;margin:18mm}body{font-family:"Microsoft YaHei",sans-serif;font-size:11pt;color:#222}h2{text-align:center}td,th{border:.2mm solid #888;padding:2mm 3mm}table{border-collapse:collapse;width:100%}</style><h2>行程单（样例）</h2><p>申请日期：2026-09-30　行程人：青岚设计工作室　共 3 笔行程，合计 87.40 元</p><table><tr><th>序号</th><th>上车时间</th><th>起点</th><th>终点</th><th>金额（元）</th></tr><tr><td>1</td><td>09-03 08:42</td><td>海棠路</td><td>梧桐里</td><td>23.10</td></tr><tr><td>2</td><td>09-11 19:05</td><td>梧桐里</td><td>半山</td><td>41.80</td></tr><tr><td>3</td><td>09-24 13:20</td><td>半山</td><td>海棠路</td><td>22.50</td></tr></table><p>本页为样例，不是发票。</p>`);
  const f = uniq('青柚出行-行程单.pdf'); await page.pdf({ path: join(OUT, f), format: 'A4', printBackground: true }); truth.push({ file: f, kind: 'not-invoice', note: '行程单，不是发票；里面也有金额，工具不该把它算进去' }); }
await browser.close();

// 标准答案和发票放在同一个文件夹里（发票整理助手只认 PDF、OFD 和图片，不会去读这个 json）
const real = truth.filter((t) => t.kind === 'invoice'); const sum = real.reduce((a, t) => a + Math.round(t.total * 100), 0) / 100;
writeFileSync(join(OUT, '标准答案.json'), JSON.stringify({ made: '2026-10-10', seed: 20261010, note: '全部是自造的样例，不是真发票', files: truth.length, invoices: real.length, duplicates: 3, imageOnly: 1, notInvoice: 1, total: sum, items: truth }, null, 1));
console.log(`${truth.length} 个文件 → ${OUT}`); console.log(`不重样的发票 ${real.length} 张，价税合计加起来 ${sum.toFixed(2)} 元`);
