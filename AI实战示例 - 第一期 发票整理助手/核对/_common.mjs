// 几个核对脚本共用的小工具：找浏览器、读标准答案、起一个只有本机能访问的临时网页服务。
//
// 依赖：Node.js 18 或更新的版本，加 playwright。在「核对」这个文件夹里运行：
//     npm install playwright
//     npx playwright install chromium
// 已经在别处装过 playwright 的，用环境变量 PLAYWRIGHT_PATH 指过去；想用电脑上现成的 Chrome / Edge，用 CHROME_PATH 指向它的 exe。
import { createRequire } from 'node:module';
import { createServer } from 'node:http';
import { existsSync, mkdirSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
export const HERE = dirname(fileURLToPath(import.meta.url));
export const SAMPLES = join(HERE, '..', '样例发票');            // 85 个样例 PDF 和标准答案.json 都在这里
export const sampleNames = () => readdirSync(SAMPLES).filter((n) => n.toLowerCase().endsWith('.pdf'));
export const readTruth = () => JSON.parse(readFileSync(join(SAMPLES, '标准答案.json'), 'utf8'));
export const outDir = () => { const d = join(HERE, '输出'); mkdirSync(d, { recursive: true }); return d; };   // 截图、导出的表格和压缩包放这里

// 命令行里第一个不带 -- 的参数就是要测的那一版 html；没给就报用法
export function htmlArg(usage) {
  const f = process.argv.slice(2).find((a) => !a.startsWith('--'));
  if (!f || !existsSync(resolve(f))) { console.error(`用法：${usage}`); process.exit(1); }
  return resolve(f);
}

export function launch() {
  let chromium;
  for (const name of [process.env.PLAYWRIGHT_PATH, 'playwright', 'playwright-core'].filter(Boolean)) { try { chromium = require(name).chromium; break; } catch { /* 试下一个 */ } }
  if (!chromium) { console.error('没找到 playwright。先在这个文件夹里运行：npm install playwright，再运行：npx playwright install chromium'); process.exit(1); }
  return chromium.launch(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {});
}

// 起一个只有本机能访问的临时网页服务（端口由系统挑）：/app.html 是要测的那一版，/发票/<文件名> 是样例发票。
// 返回 { origin, isLocal(url), close() }
export function serve(htmlFile) {
  const server = createServer((req, res) => {
    const p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (p === '/app.html') { res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }); res.end(readFileSync(htmlFile)); return; }
    const n = p.startsWith('/发票/') ? p.slice(4) : '';
    if (n && !/[\\/]/.test(n) && n.toLowerCase().endsWith('.pdf') && existsSync(join(SAMPLES, n))) { res.writeHead(200, { 'content-type': 'application/pdf' }); res.end(readFileSync(join(SAMPLES, n))); return; }
    res.writeHead(404); res.end('not found');
  });
  return new Promise((ok) => server.listen(0, '127.0.0.1', () => {
    const origin = `http://127.0.0.1:${server.address().port}`;
    ok({ origin, isLocal: (u) => String(u).startsWith(origin), close: () => server.close() });
  }));
}
