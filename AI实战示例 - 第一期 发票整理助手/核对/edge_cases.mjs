// 把「发票整理助手.html」里认发票的那段核心逻辑单独拿出来，喂几种样例里没有的情况，看它怎么判。
// 这里的发票文字是我们自己照常见写法拟的，不是真发票。
// 用法：node edge_cases.mjs <某一版.html>      （只要 Node.js，不用装别的；每一版都能测）
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
if (!process.argv[2]) { console.error('用法：node edge_cases.mjs <某一版.html>'); process.exit(1); }
const html = readFileSync(process.argv[2], 'utf8');
const core = html.match(/<script id="core">([\s\S]*?)<\/script>/)[1];
const ctx = vm.createContext({ console, TextEncoder, Blob, crypto: globalThis.crypto, document: {}, globalThis: {} });
vm.runInContext(core + '\n;globalThis.__x = { parseInvoiceText, classify, analyze, makeName };', ctx);
const { parseInvoiceText, analyze } = ctx.globalThis.__x;
const show = (name, text) => { const r = parseInvoiceText(text); console.log(`\n■ ${name}\n  kind=${r.kind} 日期=${r.date || '∅'} 销售方=${r.seller || '∅'} 购买方=${r.buyer || '∅'} 合计=${r.amount ?? '∅'} 不含税=${r.net ?? '∅'} 税=${r.tax ?? '∅'} 号码=${r.number || '∅'} 代码=${r.code || '∅'} 项目=${r.item || '∅'}`); return r; };

show('多行项目 + 合计行 + 金额带逗号', `电子发票（普通发票）\n发票号码：26112000000111111111\n开票日期：2026年09月05日\n购买方信息 销售方信息\n名称：青岚设计工作室 名称：北辰办公家具有限公司\n项目名称 规格型号 单位 数量 单价 金额 税率/征收率 税额\n*家具*办公椅 A1 把 10 1061.95 10619.47 13% 1380.53\n*家具*办公桌 B2 张 2 884.96 1769.91 13% 230.09\n合 计 ¥12,389.38 ¥1,610.62\n价税合计（大写） ⊗壹万肆仟圆整 （小写）¥14,000.00`);
show('红字发票（负数）', `电子发票（普通发票）\n发票号码：26112000000222222222\n开票日期：2026年09月06日\n名称：青岚设计工作室 名称：海棠路小馆餐饮店\n*餐饮服务*餐费 -99.01 1% -0.99\n合 计 ¥-99.01 ¥-0.99\n价税合计（大写） （负数）壹佰圆整 （小写）¥-100.00`);
show('老式电子普票（发票代码 12 位 + 号码 8 位）', `某某增值税电子普通发票\n发票代码:011002100311\n发票号码:12345678\n开票日期:2021年03月08日\n校验码:12345 67890 12345 67890\n购 名称:青岚设计工作室\n买 纳税人识别号:92SAMPLE0000000001\n货物或应税劳务、服务名称 规格型号 单位 数量 单价 金额 税率 税额\n*餐饮服务*餐费 1 94.34 94.34 6% 5.66\n合 计 ¥94.34 ¥5.66\n价税合计(大写) 壹佰圆整 (小写)¥100.00\n销 名称:海棠路小馆餐饮店\n售 纳税人识别号:92SAMPLE0000000012`);
show('没有销售方的票（火车票那一类，文字是照公开样式拟的）', `电子发票（铁路电子客票）\n发票号码:26119110000000000001\n开票日期:2026年09月12日\n北京南 G101 上海虹桥\n2026年09月10日 09:00开 05车07F号 二等座\n票价:￥553.00\n购买方名称:青岚设计工作室 统一社会信用代码:92SAMPLE0000000001`);
show('免税（税额栏是 ***）', `电子发票（普通发票）\n发票号码：26112000000333333333\n开票日期：2026年09月07日\n名称：青岚设计工作室 名称：向阳加油站\n*汽油*92号车用汽油 升 30 7.00 210.00 免税 ***\n合 计 ¥210.00 ***\n价税合计（大写） 贰佰壹拾圆整 （小写）¥210.00`);
show('销售方名字里有空格和括号', `电子发票（普通发票）\n发票号码：26112000000444444444\n开票日期：2026年09月08日\n名称：青岚设计工作室 名称：ABC Trading (Shanghai) Co., Ltd.\n*信息技术服务*软件服务费 1 100.00 100.00 6% 6.00\n价税合计（大写） 壹佰零陆圆整 （小写）¥106.00`);
show('销售方名字太长，折成了两行', `电子发票（普通发票）\n发票号码：26112000000555555555\n开票日期：2026年09月09日\n名称：青岚设计工作室 名称：云栖客栈酒店管理有限公司杭州西湖区\n统一社会信用代码/纳税人识别号： 灵隐路第二分公司\n*住宿服务*住宿费 1 300.00 300.00 6% 18.00\n价税合计（大写） 叁佰壹拾捌圆整 （小写）¥318.00`);

// 两张不同的发票，日期、销售方、金额都一样：名字会不会撞
const mk = (name, number) => ({ name, path: name, hash: name, kind: 'invoice', date: '2026-09-01', seller: '半山咖啡馆', amount: 30, number, code: '', category: '餐饮', manual: false });
const two = analyze([mk('a.pdf', '1'.repeat(20)), mk('b.pdf', '2'.repeat(20)), mk('c.pdf', '')], { nameFmt: 'date_seller_amount', groupBy: 'category' });
console.log('\n■ 同一天同一家同样金额的三张（两张号码不同，一张没读到号码）'); for (const t of two) console.log(`  ${t.name} → ${t.status}  ${t.folder}/${t.newName}`);
