/* ============================================================================
 * Tool tổng hợp PO — logic ứng dụng (chạy offline, không gửi dữ liệu ra ngoài)
 * ==========================================================================*/
(function () {
'use strict';
var C = window.NTSConvert, X = window.XLSX;
var DAY = 86400000, EMPTY = '(Trống)';

/* ----------------------------- tiện ích ----------------------------- */
function $(s, el) { return (el || document).querySelector(s); }
function $$(s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); }
function esc(s) { return s == null ? '' : String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function fmtN(n, d) { if (n == null || isNaN(n)) return '–'; return Number(n).toLocaleString('vi-VN', { minimumFractionDigits: d || 0, maximumFractionDigits: d || 0 }); }
function pct(a, b, d) { return b ? fmtN(a * 100 / b, d == null ? 1 : d) + '%' : '–'; }
function p2(n) { return String(n).padStart(2, '0'); }
function fmtD(t) { if (t == null) return ''; var d = new Date(t); return p2(d.getUTCDate()) + '/' + p2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear(); }
function fmtDT(t) { if (t == null) return ''; var d = new Date(t); return fmtD(t) + ' ' + p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()); }
function isoD(t) { var d = new Date(t); return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate()); }
function parseIso(s) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : null; }
function dayStart(t) { return Math.floor(t / DAY) * DAY; }
function monthStart(y, m) { return Date.UTC(y, m, 1); }
function monthEnd(y, m) { return Date.UTC(y, m + 1, 1) - DAY; }
function ymLabel(ym) { if (!ym || ym === EMPTY) return EMPTY; var p = ym.split('-'); return 'T' + (+p[1]) + '/' + p[0]; }
function nowNaive() { var n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate(), n.getHours(), n.getMinutes()); }
function nk(x) { return String(x == null ? '' : x).toLowerCase().replace(/đ/g, 'd').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim(); }
function median(a) { if (!a.length) return null; var s = a.slice().sort(function (x, y) { return x - y; }), m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; }
function mean(a) { if (!a.length) return null; return a.reduce(function (x, y) { return x + y; }, 0) / a.length; }
function debounce(fn, ms) { var t; return function () { var a = arguments, s = this; clearTimeout(t); t = setTimeout(function () { fn.apply(s, a); }, ms); }; }
function countBy(rows, f) { var m = new Map(); rows.forEach(function (r) { var v = f(r); (Array.isArray(v) ? v : [v]).forEach(function (k) { m.set(k, (m.get(k) || 0) + 1); }); }); return m; }
function sortedEntries(m) { return Array.from(m.entries()).sort(function (a, b) { return b[1] - a[1] || String(a[0]).localeCompare(String(b[0]), 'vi'); }); }
var store = {
  get: function (k, d) { try { var v = localStorage.getItem('ntsPO.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set: function (k, v) { try { localStorage.setItem('ntsPO.' + k, JSON.stringify(v)); } catch (e) { } }
};
function toast(msg, err) {
  var t = document.createElement('div'); t.className = 'toast' + (err ? ' err' : '');
  t.innerHTML = '<svg class="i"><use href="#i-' + (err ? 'alert' : 'check') + '"/></svg>' + esc(msg);
  document.body.appendChild(t); setTimeout(function () { t.remove(); }, err ? 5000 : 2600);
}
function loading(on, msg) {
  var el = $('#loadingEl');
  if (!on) { if (el) el.remove(); return; }
  if (!el) { el = document.createElement('div'); el.id = 'loadingEl'; el.className = 'loading'; document.body.appendChild(el); }
  el.innerHTML = '<div style="text-align:center"><div class="spinner"></div>' + esc(msg || 'Đang xử lý…') + '</div>';
}
function icon(n) { return '<svg class="i"><use href="#i-' + n + '"/></svg>'; }
function download(bytes, name, mime) {
  var blob = new Blob([bytes], { type: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

/* ----------------------------- cấu hình ----------------------------- */
var DEFAULT_RULES = [
  'An ninh – Quốc phòng | công an, catp, ca tỉnh, ca tp, bộ công an, quân khu, quân đội, quân chủng, bộ quốc phòng, bộ tư lệnh, cơ yếu, an ninh, cảnh sát, biên phòng, binh chủng, sư đoàn, lữ đoàn, trung đoàn, c06, a05, k02',
  'Tài chính – Ngân hàng – Bảo hiểm | ngân hàng, bank, chứng khoán, securities, bảo hiểm, insurance, tài chính, finance, financial, quỹ đầu tư, tín dụng, credit, depository, kho bạc, clearing',
  'Y tế | bệnh viện, hospital, y tế, dược phẩm, pharma, medical, phòng khám, clinic, y học, healthcare',
  'Giáo dục | đại học, học viện, cao đẳng, university, college, academy, school, trường thpt, trường thcs, trường tiểu học, mầm non, giáo dục, education, trung cấp',
  'Năng lượng – Điện – Dầu khí | điện lực, nhiệt điện, thủy điện, thuỷ điện, evn, genco, power, năng lượng, energy, dầu khí, petro, petrochemical, xăng dầu, petrolimex, khoáng sản, điện gió, điện mặt trời, truyền tải điện',
  'Cơ quan Nhà nước | ^sở, ^bộ, ^cục, ubnd, ủy ban nhân dân, uỷ ban nhân dân, ủy ban, uỷ ban, people\'s committee, ministry, tổng cục, chi cục, cục, ban quản lý, management board, ban duy tu, hải quan, thuế, tòa án, toà án, viện kiểm sát, tỉnh ủy, thành ủy, văn phòng chính phủ, quốc hội, chính phủ',
  'Viễn thông – CNTT | viễn thông, telecom, telecommunication, vnpt, viettel, mobifone, fpt, cmc, công nghệ thông tin, cntt, information technology, software, phần mềm, data center, chuyển đổi số, digital',
  'Doanh nghiệp | công ty, cty, tnhh, cổ phần, jsc, company, co., ltd, limited, corporation, corp, tập đoàn, group, enterprise, chi nhánh, inc',
  'Đơn vị sự nghiệp / Khác | trung tâm, center, centre, viện, institute, ban, văn phòng, hội, hiệp hội'
].join('\n');
var CFG_DEF = { m1: 7, m2: 14, m3: 30, giaoSoon: 7, licSoon: 90, rules: DEFAULT_RULES };
var CFG = Object.assign({}, CFG_DEF, store.get('cfg', {}));
var RULES = [];
function compileRules() {
  RULES = [];
  String(CFG.rules || '').split('\n').forEach(function (line) {
    var i = line.indexOf('|'); if (i < 0) return;
    var g = line.slice(0, i).trim(), kws = line.slice(i + 1).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    if (!g || !kws.length) return;
    var parts = kws.map(function (k) {
      var start = k[0] === '^'; if (start) k = k.slice(1);
      var e = nk(k).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      return (start ? '^' : '(?:^|[^a-z0-9])') + e + '(?=$|[^a-z0-9])';
    });
    RULES.push({ g: g, re: new RegExp(parts.join('|')) });
  });
}
function classify(name) {
  if (!name || name === EMPTY) return 'Chưa xác định';
  var s = nk(name);
  for (var i = 0; i < RULES.length; i++) if (RULES[i].re.test(s)) return RULES[i].g;
  return 'Chưa xác định';
}

var FLAG_META = {
  ok: { c: '#16A34A', b: 'b-ok' }, m1: { c: '#EAB308', b: 'b-yellow' }, m2: { c: '#F97316', b: 'b-orange' }, m3: { c: '#DC2626', b: 'b-red' }, done: { c: '#94A3B8', b: 'b-gray' }
};
function flagLabels() {
  return {
    ok: 'Bình thường (≤ ' + CFG.m1 + ' ngày)', m1: 'Mức 1 · ' + (CFG.m1 + 1) + '–' + CFG.m2 + ' ngày', m2: 'Mức 2 · ' + (CFG.m2 + 1) + '–' + CFG.m3 + ' ngày',
    m3: 'Mức 3 · > ' + CFG.m3 + ' ngày', done: 'Đã hoàn thành'
  };
}
var GIAO_META = { late: ['Trễ hạn giao', '#DC2626', 'b-red'], soon: ['', '#F59E0B', 'b-warn'], ok: ['', '#0EA5E9', 'b-blue'], none: ['Chưa có ngày giao', '#94A3B8', 'b-gray'], done: ['Đã hoàn thành', '#16A34A', 'b-ok'] };
function giaoLabel(k) { return k === 'soon' ? 'Đến hạn ≤ ' + CFG.giaoSoon + ' ngày' : k === 'ok' ? 'Còn > ' + CFG.giaoSoon + ' ngày' : GIAO_META[k][0]; }
var LIC_META = { expired: ['Đã hết hạn', '#64748B', 'b-gray'], soon: ['', '#DC2626', 'b-red'], active: ['', '#16A34A', 'b-ok'], none: ['Chưa có ngày', '#CBD5E1', 'b-gray'] };
function licLabel(k) { return k === 'soon' ? 'Hết hạn ≤ ' + CFG.licSoon + ' ngày' : k === 'active' ? 'Còn hạn > ' + CFG.licSoon + ' ngày' : LIC_META[k][0]; }

var ST_COLOR = { 'Hoàn thành': '#16A34A', 'Đang xử lý': '#F59E0B' };
var FLOW_COLOR = { 'Trực tiếp': '#0959A2', 'Exwork': '#F06723' };
var HD_COLOR = { 'Đã ký': '#16A34A', 'Đã ký và bàn giao cho kế toán': '#0EA5E9', 'KH đã ký nhưng chưa trả Hợp đồng': '#F59E0B', 'Chưa ký': '#DC2626' };
var PALETTE = ['#0959A2', '#F06723', '#00A8E8', '#14B8A6', '#7C3AED', '#EAB308', '#EC4899', '#22C55E', '#0EA5E9', '#F43F5E', '#8B5CF6', '#64748B', '#A3E635', '#FB923C', '#06B6D4', '#9CA3AF'];
var STEP_ORDER = ['Gửi yêu cầu đặt hàng', 'Tự động tìm PM duyệt', 'PM tiếp nhận', 'Kế toán duyệt', 'XNK tiếp nhận', 'Đặt hàng và cập nhật tiến độ', 'Hoàn thiện chứng từ', 'Kết thúc'];
var colorCache = {};
function catColor(k, i) { if (k === EMPTY) return '#CBD5E1'; if (!colorCache[k]) colorCache[k] = PALETTE[(i == null ? Object.keys(colorCache).length : i) % PALETTE.length]; return colorCache[k]; }

/* ----------------------------- trạng thái ----------------------------- */
var S = {
  fileName: '', mode: '', template: [], rows: [], warnings: [], info: '',
  refMode: 'data', refCustom: null, ref: 0, refDay: 0, filtered: [],
  sort: store.get('sort', { k: 'dTao', d: -1 }), page: 1, pageSize: 50,
  density: store.get('density', 'normal'), hangMode: 'st', calMode: 'giao'
};
var F = { q: '', facets: {}, dates: {}, quick: new Set() };

/* ----------------------------- định nghĩa bộ lọc ----------------------------- */
var FACETS = [
  { k: 'hang', label: 'Hãng', g: 0 }, { k: 'sale', label: 'Sale (chủ sở hữu)', g: 0 }, { k: 'kh', label: 'Khách hàng / Công ty', g: 0 },
  { k: 'dl', label: 'Đại lý / Reseller', g: 0 }, { k: 'nhom', label: 'Nhóm KH', g: 0 }, { k: 'pm', label: 'Người thực hiện (PM)', g: 0 }, { k: 'creator', label: 'Người tạo', g: 0 },
  { k: 'st', label: 'Trạng thái', g: 1 }, { k: 'step', label: 'Bước hiện tại', g: 1, order: STEP_ORDER }, { k: 'hd', label: 'Tình trạng HĐ với KH', g: 1 },
  { k: 'bb', label: 'BB nghiệm thu/thanh lý', g: 1 }, { k: 'flow', label: 'Luồng nhập hàng', g: 1 }, { k: 'cty', label: 'Công ty ký HĐ', g: 1 },
  { k: 'flagLabel', label: 'Cờ nhắc cập nhật', g: 1, fixed: true }, { k: 'giaoLabel', label: 'Tình trạng giao hàng', g: 1, fixed: true },
  { k: 'licLabel', label: 'Tình trạng license', g: 1, fixed: true }, { k: 'qh', label: 'Quá hạn SLA', g: 1 }, { k: 'reviews', label: 'Cờ review dữ liệu', g: 1 },
  { k: 'month', label: 'Tháng tạo', g: 2, sortKey: true }
];
var FACET_BY = {}; FACETS.forEach(function (f) { FACET_BY[f.k] = f; });
var DATES = [
  { k: 'dTao', label: 'Timeline · Ngày tạo', past: true }, { k: 'dGiao', label: 'Ngày giao hàng' },
  { k: 'dLs', label: 'License start' }, { k: 'dLe', label: 'License end' }, { k: 'dCn', label: 'Cập nhật lần cuối', past: true }
];
var DATE_BY = {}; DATES.forEach(function (d) { DATE_BY[d.k] = d; });
var GROUP_NAMES = ['Đối tượng', 'Tiến độ & Hợp đồng', 'Thời gian'];

var QUICK = [
  { id: 'inprog', label: 'Đang xử lý', c: '#F59E0B', fn: function (r) { return r.st === 'Đang xử lý'; } },
  { id: 'done', label: 'Hoàn thành', c: '#16A34A', fn: function (r) { return r.st === 'Hoàn thành'; } },
  { id: 'flag', label: 'Có cờ nhắc', c: '#F97316', fn: function (r) { return r.flag === 'm1' || r.flag === 'm2' || r.flag === 'm3'; } },
  { id: 'm3', label: 'Tồn đọng mức 3', c: '#DC2626', fn: function (r) { return r.flag === 'm3'; } },
  { id: 'late', label: 'Trễ hạn giao hàng', c: '#DC2626', fn: function (r) { return r.giao === 'late'; } },
  { id: 'soon', label: 'Sắp đến hạn giao', c: '#EAB308', fn: function (r) { return r.giao === 'soon'; } },
  { id: 'sla', label: 'Quá hạn SLA', c: '#7C3AED', fn: function (r) { return r.qh === 'Có'; } },
  { id: 'hdno', label: 'HĐ chưa ký', c: '#DC2626', fn: function (r) { return r.hd === 'Chưa ký'; } },
  { id: 'hdret', label: 'KH chưa trả HĐ', c: '#F59E0B', fn: function (r) { return r.hd === 'KH đã ký nhưng chưa trả Hợp đồng'; } },
  { id: 'gap', label: 'Hoàn thành – HĐ chưa đủ', c: '#BE123C', fn: function (r) { return r.st === 'Hoàn thành' && !r.hdOk; } },
  { id: 'nobb', label: 'Hoàn thành – chưa có BB thanh lý', c: '#64748B', fn: function (r) { return r.st === 'Hoàn thành' && r.bb !== 'Có'; } },
  { id: 'exw', label: 'Exwork', c: '#F06723', fn: function (r) { return r.flow === 'Exwork'; } },
  { id: 'lic', label: 'License sắp hết hạn', c: '#0EA5E9', fn: function (r) { return r.lic === 'soon'; } },
  { id: 'nodl', label: 'Thiếu đại lý', c: '#94A3B8', fn: function (r) { return r.dl === EMPTY; } },
  { id: 'review', label: 'Cần review dữ liệu', c: '#8B5CF6', fn: function (r) { return !!r.review; } }
];
var QUICK_BY = {}; QUICK.forEach(function (q) { QUICK_BY[q.id] = q; });

/* ----------------------------- cột bảng ----------------------------- */
function bdg(text, cls) { return text == null || text === '' ? '' : '<span class="bdg ' + (cls || '') + '"><i></i>' + esc(text) + '</span>'; }
function stBadge(v) { return bdg(v, v === 'Hoàn thành' ? 'b-ok' : v === 'Đang xử lý' ? 'b-warn' : 'b-gray'); }
function hdBadge(v) { return bdg(v, v === 'Chưa ký' ? 'b-red' : v === 'KH đã ký nhưng chưa trả Hợp đồng' ? 'b-warn' : v === 'Đã ký và bàn giao cho kế toán' ? 'b-blue' : v === 'Đã ký' ? 'b-ok' : 'b-gray'); }
var COLS = [
  { k: 'id', label: 'ID phiếu', def: 1, n: 1, html: function (r) { return '<span class="idl">#' + esc(r.id) + '</span>'; }, x: function (r) { return r.id; } },
  { k: 'no', label: 'No. (gốc)', n: 1 },
  { k: 'name', label: 'Tên phiếu', def: 1, w: 1 },
  { k: 'hang', label: 'Hãng', def: 1 },
  { k: 'kh', label: 'Khách hàng (chuẩn hóa)', def: 1 },
  { k: 'nhom', label: 'Nhóm KH', def: 1 },
  { k: 'khSrc', label: 'Nguồn tên KH' },
  { k: 'dl', label: 'Đại lý / Reseller', def: 1 },
  { k: 'sale', label: 'Sale', def: 1 },
  { k: 'pm', label: 'Người thực hiện (PM)', def: 1 },
  { k: 'creator', label: 'Người tạo' },
  { k: 'flow', label: 'Luồng nhập hàng', def: 1, html: function (r) { return bdg(r.flow, r.flow === 'Exwork' ? 'b-orange' : 'b-blue'); } },
  { k: 'hd', label: 'Tình trạng HĐ với KH', def: 1, html: function (r) { return hdBadge(r.hd); } },
  { k: 'cty', label: 'Công ty ký HĐ' },
  { k: 'bb', label: 'BB nghiệm thu/thanh lý', html: function (r) { return bdg(r.bb, r.bb === 'Có' ? 'b-ok' : r.bb === 'Không' ? 'b-gray' : ''); } },
  { k: 'dTao', label: 'Ngày tạo', def: 1, dt: 1 },
  { k: 'month', label: 'Tháng tạo', x: function (r) { return r.month; }, html: function (r) { return esc(ymLabel(r.month)); } },
  { k: 'dCn', label: 'Cập nhật lần cuối', def: 1, dt: 1 },
  { k: 'idle', label: 'Số ngày chưa cập nhật', def: 1, n: 1 },
  { k: 'flag', label: 'Cờ nhắc', def: 1, html: function (r) { return bdg(r.flagLabel, FLAG_META[r.flag].b); }, x: function (r) { return r.flagLabel; }, sv: function (r) { return { done: -1, ok: 0, m1: 1, m2: 2, m3: 3 }[r.flag]; } },
  { k: 'dGiao', label: 'Ngày giao hàng', def: 1, d: 1 },
  { k: 'daysToGiao', label: 'Số ngày đến hạn giao', n: 1 },
  { k: 'giaoLabel', label: 'Tình trạng giao hàng', html: function (r) { return bdg(r.giaoLabel, GIAO_META[r.giao][2]); } },
  { k: 'dLs', label: 'License start', d: 1 },
  { k: 'dLe', label: 'License end', def: 1, d: 1 },
  { k: 'licMonths', label: 'Thời hạn license (tháng)', n: 1, dec: 1 },
  { k: 'daysToLe', label: 'Số ngày đến hết hạn license', n: 1 },
  { k: 'licLabel', label: 'Tình trạng license', html: function (r) { return bdg(r.licLabel, LIC_META[r.lic][2]); } },
  { k: 'step', label: 'Bước hiện tại', def: 1 },
  { k: 'st', label: 'Trạng thái', def: 1, html: function (r) { return stBadge(r.st); } },
  { k: 'tg', label: 'TG xử lý (ngày, phiếu HT)', def: 1, n: 1, dec: 1 },
  { k: 'age', label: 'Tuổi phiếu (ngày)', n: 1 },
  { k: 'dSla', label: 'SLA/Hạn', dt: 1 },
  { k: 'qh', label: 'Quá hạn SLA', html: function (r) { return r.qh === 'Có' ? bdg('Quá hạn', 'b-purple') : esc(r.qh === EMPTY ? '' : r.qh); } },
  { k: 'hdOkLabel', label: 'Tuân thủ HĐ', html: function (r) { return bdg(r.hdOkLabel, r.hdOkLabel === 'Đạt' ? 'b-ok' : r.hdOkLabel === 'Chưa đạt' ? 'b-red' : 'b-gray'); } },
  { k: 'review', label: 'Cờ review', w: 1 },
  { k: 'action', label: 'Đề xuất hành động', w: 1 },
  { k: 'mota', label: 'Mô tả phiếu', w: 1 },
  { k: 'bom', label: 'BOM', w: 1 }
];
var COL_BY = {}; COLS.forEach(function (c) { COL_BY[c.k] = c; });
var COL_PRESETS = {
  'Mặc định': COLS.filter(function (c) { return c.def; }).map(function (c) { return c.k; }),
  'Tối giản': ['id', 'name', 'hang', 'kh', 'sale', 'st', 'flag'],
  'Theo dõi tiến độ': ['id', 'name', 'hang', 'sale', 'pm', 'step', 'st', 'dTao', 'dCn', 'idle', 'flag', 'dGiao', 'giaoLabel', 'qh', 'action'],
  'Hợp đồng & License': ['id', 'name', 'hang', 'kh', 'dl', 'sale', 'hd', 'cty', 'bb', 'hdOkLabel', 'dLs', 'dLe', 'licMonths', 'daysToLe', 'licLabel'],
  'Tất cả': COLS.map(function (c) { return c.k; })
};
var visCols = store.get('cols', null);
if (!Array.isArray(visCols) || !visCols.length) visCols = COL_PRESETS['Mặc định'].slice();
visCols = visCols.filter(function (k) { return COL_BY[k]; });

function cellText(r, c) {
  var v = r[c.k];
  if (c.x) return c.x(r);
  if (c.dt) return fmtDT(v); if (c.d) return fmtD(v);
  if (v === EMPTY) return '';
  return v;
}
function cellHtml(r, c) {
  if (c.html) return c.html(r);
  var v = r[c.k];
  if (c.dt) return esc(fmtDT(v)); if (c.d) return esc(fmtD(v));
  if (c.n) return v == null ? '' : esc(fmtN(v, c.dec ? 1 : 0));
  if (v === EMPTY || v == null) return '<span style="color:var(--muted)">–</span>';
  return esc(v);
}
function sortVal(r, k) { var c = COL_BY[k]; if (c && c.sv) return c.sv(r); var v = r[k]; return v === EMPTY ? null : v; }

/* ----------------------------- nạp file ----------------------------- */
function handleFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm|xls|csv)$/i.test(file.name)) { toast('Vui lòng chọn file Excel (.xlsx/.xls)', true); return; }
  loading(true, 'Đang đọc & chuẩn hóa dữ liệu…');
  var fr = new FileReader();
  fr.onload = function (e) {
    setTimeout(function () {
      try { loadWorkbook(new Uint8Array(e.target.result), file.name); }
      catch (err) { console.error(err); loading(false); toast(err.message || String(err), true); }
    }, 30);
  };
  fr.onerror = function () { loading(false); toast('Không đọc được file', true); };
  fr.readAsArrayBuffer(file);
}
function loadWorkbook(bytes, name) {
  var wb = X.read(bytes, { type: 'array', cellDates: false });
  var mode = null, ws = null, missing = null;
  // 1) file đã chuẩn hóa (Data_sạch)?
  wb.SheetNames.some(function (n) {
    var h = (X.utils.sheet_to_json(wb.Sheets[n], { header: 1, range: 0, blankrows: false })[0] || []).map(function (x) { return String(x || '').trim(); });
    if (h.indexOf('ID phiếu') >= 0 && h.indexOf('Hãng') >= 0 && h.indexOf('Bước hiện tại') >= 0) { mode = 'clean'; ws = wb.Sheets[n]; return true; }
    if (h.indexOf('ID') >= 0 && h.indexOf('Tên Hãng') >= 0) { mode = 'raw'; ws = wb.Sheets[n]; return true; }
    return false;
  });
  var template, warnings = [];
  if (mode === 'clean') {
    template = C.readClean(X, ws);
  } else {
    ws = ws || wb.Sheets[wb.SheetNames[0]];
    var raw = C.sheetToRows(X, ws).rows;
    if (!raw.length) throw new Error('File không có dữ liệu.');
    var res = C.convert(raw);
    template = res.rows; warnings = res.warnings; mode = 'raw';
  }
  if (!template.length) throw new Error('Không tìm thấy dòng dữ liệu nào.');
  S.fileName = name; S.mode = mode; S.template = template; S.warnings = warnings;
  S.rows = template.map(enrichBase);
  colorCache = {};
  sortedEntries(countBy(S.rows, function (r) { return r.hang; })).forEach(function (e, i) { catColor(e[0], i); });
  resetFilters(true);
  derive();
  initUI();
  loading(false);
  toast((mode === 'raw' ? 'Đã convert ' : 'Đã nạp ') + template.length + ' phiếu từ “' + name + '”');
}

function enrichBase(t, i) {
  var r = { t: t, _i: i };
  function s(k) { var v = t[k]; return v == null || v === '' ? EMPTY : String(v); }
  r.id = t['ID phiếu']; r.no = t['No. (gốc)']; r.name = t['Tên phiếu'] || '';
  r.hang = s('Hãng'); r.kh = s('Khách hàng (chuẩn hóa)'); r.khSrc = s('Nguồn tên KH'); r.dl = s('Đại lý'); r.sale = s('Sale');
  r.creator = s('Người tạo'); r.pm = s('Người thực hiện');
  r.flow = s('Luồng nhập hàng'); if (r.flow === 'Exw') r.flow = 'Exwork';
  r.hd = s('Tình trạng HĐ với KH'); r.cty = s('Công ty ký HĐ (*)'); r.bb = s('Biên bản nghiệm thu/thanh lý (*)');
  r.step = s('Bước hiện tại'); r.st = s('Trạng thái'); r.qh = s('Quá hạn (*)');
  r.tg = t['Thời gian xử lý (ngày, phiếu hoàn thành)']; r.review = t['Cờ review'] || '';
  r.reviews = r.review ? r.review.split('; ') : ['(Không có)'];
  r.mota = t['Mô tả phiếu'] || ''; r.bom = t['BOM'] || '';
  r.dTao = t['Ngày tạo']; r.dCn = t['Cập nhật lần cuối']; r.dGiao = t['Ngày giao hàng']; r.dLs = t['License start (*)']; r.dLe = t['License end']; r.dSla = t['SLA/Hạn (*)'];
  r.month = t['Tháng tạo'] || (r.dTao != null ? C.ymOf(r.dTao) : EMPTY);
  r.licMonths = (r.dLs != null && r.dLe != null) ? Math.round((r.dLe - r.dLs) / DAY / 30.4375 * 10) / 10 : null;
  r.hdOk = r.hd === 'Đã ký' || r.hd === 'Đã ký và bàn giao cho kế toán';
  r.hay = nk([r.id, r.name, r.hang, r.kh, r.dl, r.sale, r.pm, r.creator, r.step, r.st, r.hd, r.flow, r.cty, r.mota, r.bom, r.review].join(' | '));
  return r;
}
function computeRef() {
  if (S.refMode === 'today') S.ref = nowNaive();
  else if (S.refMode === 'custom' && S.refCustom != null) S.ref = S.refCustom + DAY - 60000;
  else { var m = 0; S.rows.forEach(function (r) { if (r.dCn != null && r.dCn > m) m = r.dCn; }); S.ref = m || nowNaive(); }
  S.refDay = dayStart(S.ref);
}
function derive() {
  compileRules(); computeRef();
  var FL = flagLabels();
  S.rows.forEach(function (r) {
    r.nhom = classify(r.kh !== EMPTY ? r.kh : r.name);
    r.idle = r.dCn != null ? Math.max(0, Math.floor((S.ref - r.dCn) / DAY)) : null;
    r.age = r.dTao != null ? Math.max(0, Math.floor((S.ref - r.dTao) / DAY)) : null;
    var done = r.st === 'Hoàn thành';
    if (done) r.flag = 'done';
    else if (r.idle == null || r.idle > CFG.m3) r.flag = 'm3';
    else if (r.idle > CFG.m2) r.flag = 'm2';
    else if (r.idle > CFG.m1) r.flag = 'm1';
    else r.flag = 'ok';
    r.flagLabel = FL[r.flag];
    r.daysToGiao = r.dGiao != null ? Math.round((r.dGiao - S.refDay) / DAY) : null;
    r.giao = done ? 'done' : r.dGiao == null ? 'none' : r.daysToGiao < 0 ? 'late' : r.daysToGiao <= CFG.giaoSoon ? 'soon' : 'ok';
    r.giaoLabel = giaoLabel(r.giao);
    r.daysToLe = r.dLe != null ? Math.round((r.dLe - S.refDay) / DAY) : null;
    r.lic = r.dLe == null ? 'none' : r.daysToLe < 0 ? 'expired' : r.daysToLe <= CFG.licSoon ? 'soon' : 'active';
    r.licLabel = licLabel(r.lic);
    r.hdOkLabel = done ? (r.hdOk ? 'Đạt' : 'Chưa đạt') : 'Đang xử lý';
    var a = [];
    if (!done) {
      if (r.flag === 'm3') a.push('Tồn đọng ' + (r.idle == null ? '' : r.idle + ' ngày ') + 'không cập nhật – escalate PM/Sale');
      else if (r.flag === 'm2' || r.flag === 'm1') a.push('Nhắc ' + (r.pm !== EMPTY ? r.pm : 'PM') + ' cập nhật tiến độ (' + r.idle + ' ngày)');
      if (r.giao === 'late') a.push('Trễ hạn giao ' + (-r.daysToGiao) + ' ngày – xác nhận lại tiến độ với hãng/KH');
      else if (r.giao === 'soon') a.push('Sắp đến hạn giao (' + fmtD(r.dGiao) + ')');
      if (r.qh === 'Có') a.push('Quá hạn SLA ở bước “' + r.step + '”');
    } else {
      if (r.hd === 'Chưa ký') a.push('Đã hoàn thành nhưng HĐ chưa ký – nhắc Sale ký HĐ');
      else if (r.hd === 'KH đã ký nhưng chưa trả Hợp đồng') a.push('Thu hồi HĐ bản gốc từ KH');
      if (r.lic === 'soon') a.push('Chuẩn bị renew – license hết hạn ' + fmtD(r.dLe));
    }
    if (r.dl === EMPTY) a.push('Bổ sung đại lý');
    r.action = a.join('; ');
  });
  applyFilters();
}

/* ----------------------------- lọc ----------------------------- */
function resetFilters(all) {
  F.q = ''; F.facets = {}; F.dates = {}; F.quick = new Set();
  FACETS.forEach(function (f) { F.facets[f.k] = new Set(); });
  if ($('#q')) { $('#q').value = ''; }
}
function vals(r, k) { var v = r[k]; return Array.isArray(v) ? v : [v]; }
function passes(r, skip) {
  if (F.q && skip !== 'q') {
    var toks = nk(F.q).split(' ');
    for (var i = 0; i < toks.length; i++) if (toks[i] && r.hay.indexOf(toks[i]) < 0) return false;
  }
  for (var k in F.facets) {
    if (k === skip) continue;
    var set = F.facets[k]; if (!set.size) continue;
    var vs = vals(r, k), ok = false;
    for (var j = 0; j < vs.length; j++) if (set.has(vs[j])) { ok = true; break; }
    if (!ok) return false;
  }
  for (var dk in F.dates) {
    if (dk === skip) continue;
    var d = F.dates[dk]; if (!d) continue;
    var v = r[dk];
    if (d.empty) { if (v != null) return false; continue; }
    if (v == null) return false;
    if (d.from != null && v < d.from) return false;
    if (d.to != null && v > d.to + DAY - 1) return false;
  }
  var it = F.quick.values(), q;
  while (!(q = it.next()).done) { if (skip === 'quick:' + q.value) continue; if (!QUICK_BY[q.value].fn(r)) return false; }
  return true;
}
function applyFilters() {
  S.filtered = S.rows.filter(function (r) { return passes(r); });
  var k = S.sort.k, d = S.sort.d;
  S.filtered.sort(function (a, b) {
    var x = sortVal(a, k), y = sortVal(b, k);
    if (x == null && y == null) return a._i - b._i; if (x == null) return 1; if (y == null) return -1;
    if (typeof x === 'string' || typeof y === 'string') { var c = String(x).localeCompare(String(y), 'vi'); return c * d || a._i - b._i; }
    return (x - y) * d || a._i - b._i;
  });
}
function activeCount() {
  var n = (F.q ? 1 : 0) + F.quick.size;
  for (var k in F.facets) if (F.facets[k].size) n++;
  for (var d in F.dates) if (F.dates[d]) n++;
  return n;
}
function toggleFacet(k, v, only) {
  var s = F.facets[k]; if (!s) return;
  if (only) { var had = s.size === 1 && s.has(v); s.clear(); if (!had) s.add(v); }
  else if (s.has(v)) s.delete(v); else s.add(v);
  S.page = 1; update();
}
function setDate(k, val) { F.dates[k] = val; S.page = 1; update(); }

var update = function () { applyFilters(); renderAll(); };
var updateDebounced = debounce(update, 160);

/* ----------------------------- UI init ----------------------------- */
var uiReady = false;
function initUI() {
  $('#heroEmpty').classList.add('hidden'); $('#heroData').classList.remove('hidden');
  $('#app').classList.remove('hidden');
  ['#btnReload', '#btnClean', '#btnExport', '#tbFile'].forEach(function (s) { $(s).classList.remove('hidden'); });
  $('#btnClean').classList.toggle('hidden', false);
  $('#tbFileName').textContent = S.fileName;
  if (!uiReady) { buildFilterUI(); bindStatic(); uiReady = true; }
  $('#refMode').value = S.refMode;
  renderWarnings();
  renderAll();
}
function buildFilterUI() {
  var qh = '<span class="lbl">Lọc nhanh</span>';
  QUICK.forEach(function (q) { qh += '<span class="qc" data-q="' + q.id + '"><i class="d" style="background:' + q.c + '"></i>' + esc(q.label) + '<span class="n">0</span></span>'; });
  $('#quick').innerHTML = qh;
  var gh = '';
  GROUP_NAMES.forEach(function (gn, gi) {
    gh += '<div class="f-group"><span class="lbl">' + gn + '</span>';
    FACETS.filter(function (f) { return f.g === gi; }).forEach(function (f) {
      if (gi === 2) return;
      gh += '<div class="fp" data-fk="' + f.k + '"><button type="button"><span class="lb">' + esc(f.label) + '</span><span class="c hidden"></span><svg class="i chev"><use href="#i-chev"/></svg></button></div>';
    });
    if (gi === 2) {
      DATES.forEach(function (d) { gh += '<div class="fp" data-dk="' + d.k + '"><button type="button"><svg class="i"><use href="#i-cal"/></svg><span class="lb">' + esc(d.label) + '</span><svg class="i chev"><use href="#i-chev"/></svg></button></div>'; });
      gh += '<div class="fp" data-fk="month"><button type="button"><span class="lb">Tháng tạo</span><span class="c hidden"></span><svg class="i chev"><use href="#i-chev"/></svg></button></div>';
    }
    gh += '</div>';
  });
  $('#fgroups').innerHTML = gh;
}

/* ----------------------------- render tổng ----------------------------- */
function renderAll() {
  $('#tbCount').textContent = '· ' + fmtN(S.filtered.length) + '/' + fmtN(S.rows.length) + ' phiếu';
  renderRef(); renderHero(); renderQuick(); renderPills(); renderActive(); renderAlert();
  renderKPIs(); renderOverview(); renderInsights(); renderCharts(); renderStatTables(); renderTable();
  refreshOpenPop();
}
function renderRef() {
  $('#refLabel').textContent = fmtDT(S.ref);
  $('#refDate').classList.toggle('hidden', S.refMode !== 'custom');
}
function periodOf(rows) {
  var mn = null, mx = null;
  rows.forEach(function (r) { if (r.dTao != null) { if (mn == null || r.dTao < mn) mn = r.dTao; if (mx == null || r.dTao > mx) mx = r.dTao; } });
  return [mn, mx];
}
function renderHero() {
  var p = periodOf(S.rows), R = S.filtered;
  $('#heroPeriod').textContent = 'Tháng ' + (new Date(S.ref).getUTCMonth() + 1) + '.' + new Date(S.ref).getUTCFullYear();
  $('#heroSub').innerHTML = 'Nguồn: <b>' + esc(S.fileName) + '</b> · ' + (S.mode === 'raw' ? 'đã convert tự động từ file raw Base Service sang template Data_sạch' : 'file template Data_sạch') +
    ' · Dữ liệu tạo phiếu từ <b>' + fmtD(p[0]) + '</b> đến <b>' + fmtD(p[1]) + '</b>';
  var nH = new Set(R.map(function (r) { return r.hang; })).size, nS = new Set(R.map(function (r) { return r.sale; })).size, nK = new Set(R.map(function (r) { return r.kh; })).size, nD = new Set(R.filter(function (r) { return r.dl !== EMPTY; }).map(function (r) { return r.dl; })).size;
  $('#heroMeta').innerHTML = [['Phiếu', fmtN(R.length) + ' / ' + fmtN(S.rows.length)], ['Hãng', nH], ['Sale', nS], ['Khách hàng', nK], ['Đại lý', nD], ['Đối chiếu', fmtDT(S.ref)]]
    .map(function (x) { return '<span class="pill-glass">' + x[0] + ': <b>' + x[1] + '</b></span>'; }).join('');
}
function renderWarnings() {
  var h = '';
  if (S.warnings.length) h = '<div class="warnbox"><b>Cảnh báo khi convert</b> (dữ liệu vẫn được giữ nguyên, nên rà soát mapping):<ul>' + S.warnings.map(function (w) { return '<li>' + esc(w) + '</li>'; }).join('') + '</ul></div>';
  $('#warnHost').innerHTML = h;
}
function renderQuick() {
  $$('#quick .qc').forEach(function (el) {
    var id = el.dataset.q, q = QUICK_BY[id], on = F.quick.has(id);
    var n = 0; S.rows.forEach(function (r) { if (q.fn(r) && passes(r, 'quick:' + id)) n++; });
    el.classList.toggle('on', on); el.classList.toggle('zero', !n && !on);
    $('.n', el).textContent = fmtN(n);
  });
}
function renderPills() {
  $$('#fgroups .fp[data-fk]').forEach(function (el) {
    var k = el.dataset.fk, s = F.facets[k], c = $('.c', el);
    el.classList.toggle('on', s.size > 0); c.classList.toggle('hidden', !s.size); c.textContent = s.size;
  });
  $$('#fgroups .fp[data-dk]').forEach(function (el) {
    var k = el.dataset.dk, d = F.dates[k];
    el.classList.toggle('on', !!d);
    $('.lb', el).textContent = DATE_BY[k].label + (d ? ': ' + dateDesc(d) : '');
  });
}
function dateDesc(d) {
  if (d.empty) return 'Trống';
  if (d.label) return d.label;
  if (d.from != null && d.to != null) return fmtD(d.from) + ' → ' + fmtD(d.to);
  if (d.from != null) return 'từ ' + fmtD(d.from);
  if (d.to != null) return 'đến ' + fmtD(d.to);
  return '';
}
function renderActive() {
  var h = '<span class="cnt">Đang hiển thị <b>' + fmtN(S.filtered.length) + '</b> / ' + fmtN(S.rows.length) + ' phiếu</span>';
  var tags = [];
  if (F.q) tags.push(['Từ khóa', F.q, 'q']);
  F.quick.forEach(function (id) { tags.push(['Lọc nhanh', QUICK_BY[id].label, 'quick:' + id]); });
  FACETS.forEach(function (f) { var s = F.facets[f.k]; if (s.size) tags.push([f.label, Array.from(s).map(function (v) { return f.k === 'month' ? ymLabel(v) : v; }).join(', '), 'f:' + f.k]); });
  DATES.forEach(function (d) { if (F.dates[d.k]) tags.push([d.label, dateDesc(F.dates[d.k]), 'd:' + d.k]); });
  if (!tags.length) h += '<span style="font-size:12.5px;color:var(--muted)">Chưa áp dụng bộ lọc nào — có thể kết hợp nhiều bộ lọc cùng lúc.</span>';
  tags.forEach(function (t) { h += '<span class="ftag" title="' + esc(t[0] + ': ' + t[1]) + '"><span><b>' + esc(t[0]) + ':</b> ' + esc(t[1]) + '</span><button data-rm="' + esc(t[2]) + '">' + icon('x') + '</button></span>'; });
  if (tags.length > 1) h += '<button class="btn sm ghost" data-rm="*">Xóa tất cả</button>';
  $('#activeBar').innerHTML = h;
}
function removeTag(key) {
  if (key === '*') resetFilters(); else if (key === 'q') { F.q = ''; $('#q').value = ''; }
  else if (key.indexOf('quick:') === 0) F.quick.delete(key.slice(6));
  else if (key.indexOf('f:') === 0) F.facets[key.slice(2)].clear();
  else if (key.indexOf('d:') === 0) delete F.dates[key.slice(2)];
  S.page = 1; update();
}

/* ----------------------------- cảnh báo cờ nhắc ----------------------------- */
function renderAlert() {
  var R = S.filtered.filter(function (r) { return r.flag === 'm1' || r.flag === 'm2' || r.flag === 'm3'; });
  if (!R.length) { $('#alertHost').innerHTML = ''; return; }
  var old = R.slice().sort(function (a, b) { return (b.idle || 0) - (a.idle || 0); })[0];
  var FL = flagLabels(), c = countBy(R, function (r) { return r.flag; });
  var late = S.filtered.filter(function (r) { return r.giao === 'late'; }).length;
  $('#alertHost').innerHTML = '<div class="alert"><div class="ic">' + icon('bell') + '</div><div class="msg">' +
    '<b class="r">' + fmtN(R.length) + '</b> phiếu đang xử lý đã <b>hơn ' + CFG.m1 + ' ngày</b> chưa được cập nhật (tính đến ' + fmtD(S.ref) + '). Lâu nhất: <b>#' + esc(old.id) + '</b> ' + esc(old.name) + ' — <b class="r">' + fmtN(old.idle) + ' ngày</b> (PM ' + esc(old.pm) + ').' +
    (late ? ' · <b class="r">' + late + '</b> phiếu đã trễ hạn giao hàng.' : '') +
    '<div class="lv">' + ['m3', 'm2', 'm1'].map(function (k) { return '<span><i style="background:' + FLAG_META[k].c + '"></i>' + esc(FL[k]) + ' <b>' + (c.get(k) || 0) + '</b></span>'; }).join('') + '</div></div>' +
    '<button class="btn primary" id="alView">' + icon('list') + 'Xem danh sách</button><button class="btn accent" id="alMail">' + icon('mail') + 'Soạn nội dung nhắc</button></div>';
  $('#alView').onclick = function () { F.quick.add('flag'); S.sort = { k: 'idle', d: -1 }; update(); $('#tableCard').scrollIntoView({ behavior: 'smooth' }); };
  $('#alMail').onclick = openReminder;
}

/* ----------------------------- KPI ----------------------------- */
function renderKPIs() {
  var R = S.filtered, n = R.length;
  var done = R.filter(function (r) { return r.st === 'Hoàn thành'; }), ip = R.filter(function (r) { return r.st === 'Đang xử lý'; });
  var tg = done.map(function (r) { return r.tg; }).filter(function (v) { return v != null; });
  var flag = ip.filter(function (r) { return r.flag !== 'ok'; }).length;
  var late = ip.filter(function (r) { return r.giao === 'late'; }).length, sla = R.filter(function (r) { return r.qh === 'Có'; }).length;
  var hdOk = done.filter(function (r) { return r.hdOk; }).length, bb = done.filter(function (r) { return r.bb === 'Có'; }).length;
  var exw = R.filter(function (r) { return r.flow === 'Exwork'; }).length, lic = R.filter(function (r) { return r.lic === 'soon'; }).length;
  var K = [
    ['Tổng số phiếu', fmtN(n), new Set(R.map(function (r) { return r.sale; })).size + ' sale · ' + new Set(R.map(function (r) { return r.hang; })).size + ' hãng · ' + new Set(R.map(function (r) { return r.kh; })).size + ' KH', '#0959A2', 100, ''],
    ['Đang xử lý', fmtN(ip.length), pct(ip.length, n) + ' tổng số', '#F59E0B', ip.length / (n || 1) * 100, 'inprog'],
    ['Hoàn thành', fmtN(done.length), pct(done.length, n) + ' tổng số', '#16A34A', done.length / (n || 1) * 100, 'done'],
    ['TG xử lý (trung vị)', tg.length ? fmtN(median(tg), 1) : '–', 'ngày · TB ' + (tg.length ? fmtN(mean(tg), 1) : '–') + ' ngày', '#00A8E8', null, ''],
    ['Cần nhắc cập nhật', fmtN(flag), '> ' + CFG.m1 + ' ngày không cập nhật', '#F97316', flag / (ip.length || 1) * 100, 'flag'],
    ['Trễ hạn giao / SLA', fmtN(late) + ' / ' + fmtN(sla), 'trễ giao hàng · quá hạn SLA', '#DC2626', late / (ip.length || 1) * 100, 'late'],
    ['Tuân thủ ký HĐ', pct(hdOk, done.length), fmtN(hdOk) + '/' + fmtN(done.length) + ' phiếu HT đã ký HĐ', '#7C3AED', hdOk / (done.length || 1) * 100, 'gap'],
    ['BB thanh lý · Exwork', fmtN(bb) + ' · ' + fmtN(exw), 'BB “Có” trên HT · ' + fmtN(lic) + ' license sắp hết hạn', '#14B8A6', bb / (done.length || 1) * 100, 'nobb']
  ];
  $('#kpis').innerHTML = K.map(function (k) {
    return '<div class="kpi" style="--c:' + k[3] + '"' + (k[5] ? ' data-act="' + k[5] + '" title="Bấm để lọc"' : '') + '><div class="k">' + k[0] + '</div><div class="v">' + k[1] + '</div><div class="s" title="' + esc(k[2]) + '">' + esc(k[2]) + '</div>' +
      (k[4] != null ? '<div class="bar"><i style="width:' + Math.min(100, k[4]).toFixed(1) + '%"></i></div>' : '') + '</div>';
  }).join('');
}

/* ----------------------------- tổng quan donut ----------------------------- */
function legendRows(entries, total, colorOf, attr) {
  return entries.map(function (e) {
    return '<div class="row" ' + attr(e[0]) + '><i style="background:' + colorOf(e[0]) + '"></i><span>' + esc(e[0]) + '</span><b>' + fmtN(e[1]) + '</b><em>' + pct(e[1], total) + '</em></div>';
  }).join('');
}
function renderOverview() {
  var R = S.filtered, n = R.length;
  var st = sortedEntries(countBy(R, function (r) { return r.st; }));
  $('#lgStatus').innerHTML = legendRows(st, n, function (k) { return ST_COLOR[k] || '#94A3B8'; }, function (k) { return 'data-f="st" data-v="' + esc(k) + '"'; });
  var ip = R.filter(function (r) { return r.st !== 'Hoàn thành'; });
  var steps = Array.from(countBy(ip, function (r) { return r.step; }).entries()).sort(function (a, b) { return stepIdx(a[0]) - stepIdx(b[0]); });
  $('#lgSteps').innerHTML = steps.length ? steps.map(function (e, i) { return '<div class="mini" data-f="step" data-v="' + esc(e[0]) + '"><i style="background:' + PALETTE[(i + 2) % PALETTE.length] + '"></i>' + esc(e[0]) + '<b>' + e[1] + '</b></div>'; }).join('') : '<div class="mini">Không có phiếu đang xử lý</div>';
  donut('cStatus', st.map(function (e) { return e[0]; }), st.map(function (e) { return e[1]; }), st.map(function (e) { return ST_COLOR[e[0]] || '#94A3B8'; }), fmtN(n), 'PHIẾU', 'st', true);

  var FL = flagLabels(), keys = ['ok', 'm1', 'm2', 'm3'];
  var fc = countBy(ip, function (r) { return r.flag; });
  var need = (fc.get('m1') || 0) + (fc.get('m2') || 0) + (fc.get('m3') || 0);
  $('#lgFlag').innerHTML = keys.map(function (k) { var v = fc.get(k) || 0; return '<div class="row" data-f="flagLabel" data-v="' + esc(FL[k]) + '"><i style="background:' + FLAG_META[k].c + '"></i><span>' + esc(FL[k]) + '</span><b>' + v + '</b><em>' + pct(v, ip.length) + '</em></div>'; }).join('');
  donut('cFlag', keys.map(function (k) { return FL[k]; }), keys.map(function (k) { return fc.get(k) || 0; }), keys.map(function (k) { return FLAG_META[k].c; }), fmtN(need), 'CẦN NHẮC', 'flagLabel', true);
  var gc = countBy(ip, function (r) { return r.giao; });
  $('#lgGiao').innerHTML = ['late', 'soon', 'ok', 'none'].map(function (k) { return '<div class="mini" data-f="giaoLabel" data-v="' + esc(giaoLabel(k)) + '"><i style="background:' + GIAO_META[k][1] + '"></i>' + esc(giaoLabel(k)) + '<b>' + (gc.get(k) || 0) + '</b></div>'; }).join('');
}
function stepIdx(s) { var i = STEP_ORDER.indexOf(s); return i < 0 ? 50 : i; }

/* ----------------------------- insights ----------------------------- */
function renderInsights() {
  var R = S.filtered, n = R.length, L = [];
  if (!n) { $('#insights').innerHTML = '<li>Không có phiếu nào khớp bộ lọc hiện tại.</li>'; return; }
  var h = sortedEntries(countBy(R, function (r) { return r.hang; }));
  L.push('Hãng chiếm đa số: <b class="bl">' + esc(h[0][0]) + '</b> với <b>' + h[0][1] + '</b> phiếu (' + pct(h[0][1], n) + ')' + (h[1] ? ', tiếp theo <b>' + esc(h[1][0]) + '</b> ' + pct(h[1][1], n) + (h[2] ? ', <b>' + esc(h[2][0]) + '</b> ' + pct(h[2][1], n) : '') : '') + '.');
  var s = sortedEntries(countBy(R, function (r) { return r.sale; }));
  L.push('Sale nhiều phiếu nhất: <b class="bl">' + esc(s[0][0]) + '</b> (' + s[0][1] + ' phiếu, ' + pct(s[0][1], n) + '); top 3 sale chiếm ' + pct(s.slice(0, 3).reduce(function (a, e) { return a + e[1]; }, 0), n) + '.');
  var done = R.filter(function (r) { return r.st === 'Hoàn thành' && r.tg != null; });
  if (done.length) {
    var byH = {}; done.forEach(function (r) { (byH[r.hang] = byH[r.hang] || []).push(r.tg); });
    var hs = Object.keys(byH).filter(function (k) { return byH[k].length >= 3; }).map(function (k) { return [k, median(byH[k])]; }).sort(function (a, b) { return b[1] - a[1]; });
    L.push('Thời gian xử lý trung vị <b>' + fmtN(median(done.map(function (r) { return r.tg; })), 1) + ' ngày</b>' + (hs.length ? '; chậm nhất: <b>' + esc(hs[0][0]) + '</b> (' + fmtN(hs[0][1], 1) + ' ngày), nhanh nhất: <b>' + esc(hs[hs.length - 1][0]) + '</b> (' + fmtN(hs[hs.length - 1][1], 1) + ' ngày)' : '') + '.');
  }
  var ip = R.filter(function (r) { return r.st !== 'Hoàn thành'; });
  if (ip.length) {
    var stale = ip.filter(function (r) { return r.flag !== 'ok'; }), late = ip.filter(function (r) { return r.giao === 'late'; });
    var pm = sortedEntries(countBy(stale, function (r) { return r.pm; }));
    L.push('<b>' + ip.length + '</b> phiếu đang xử lý, trong đó <b class="hl">' + stale.length + '</b> cần nhắc cập nhật' + (pm.length ? ' (nhiều nhất: PM <b>' + esc(pm[0][0]) + '</b> – ' + pm[0][1] + ' phiếu)' : '') + ', <b class="hl">' + late.length + '</b> phiếu đã trễ hạn giao hàng.');
  }
  var dn = R.filter(function (r) { return r.st === 'Hoàn thành'; });
  if (dn.length) {
    var gap = dn.filter(function (r) { return !r.hdOk; }), ns = dn.filter(function (r) { return r.hd === 'Chưa ký'; }).length;
    L.push('Tuân thủ HĐ: <b>' + pct(dn.length - gap.length, dn.length) + '</b> phiếu hoàn thành đã ký HĐ; còn <b class="hl">' + ns + '</b> phiếu HĐ chưa ký và <b>' + (gap.length - ns) + '</b> phiếu KH chưa trả HĐ. BB thanh lý “Có”: <b>' + pct(dn.filter(function (r) { return r.bb === 'Có'; }).length, dn.length) + '</b>.');
  }
  var exw = R.filter(function (r) { return r.flow === 'Exwork'; }).length;
  L.push('Luồng nhập hàng: <b>' + pct(n - exw, n) + '</b> nhập trực tiếp, <b>' + pct(exw, n) + '</b> Exwork (' + exw + ' phiếu).');
  var lic = R.filter(function (r) { return r.lic === 'soon'; });
  if (lic.length) L.push('Cơ hội renew: <b class="bl">' + lic.length + '</b> phiếu có license hết hạn trong ' + CFG.licSoon + ' ngày tới (' + esc(sortedEntries(countBy(lic, function (r) { return r.hang; })).slice(0, 3).map(function (e) { return e[0] + ' ' + e[1]; }).join(', ')) + ').');
  var dl = sortedEntries(countBy(R.filter(function (r) { return r.dl !== EMPTY; }), function (r) { return r.dl; }));
  if (dl.length) L.push('Đại lý/Reseller tích cực nhất: <b class="bl">' + esc(dl[0][0]) + '</b> (' + dl[0][1] + ' phiếu); ' + R.filter(function (r) { return r.dl === EMPTY; }).length + ' phiếu chưa khai báo đại lý.');
  var ng = sortedEntries(countBy(R, function (r) { return r.nhom; }));
  L.push('Nhóm khách hàng lớn nhất: <b>' + esc(ng[0][0]) + '</b> (' + pct(ng[0][1], n) + ')' + (ng[1] ? ', <b>' + esc(ng[1][0]) + '</b> (' + pct(ng[1][1], n) + ')' : '') + '.');
  $('#insights').innerHTML = L.map(function (x) { return '<li>' + x + '</li>'; }).join('');
}

/* ----------------------------- biểu đồ (Chart.js) ----------------------------- */
var CH = {};
function cssv(n) { return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
function chartBase() {
  Chart.defaults.font.family = '"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif';
  Chart.defaults.font.size = 12;
  Chart.defaults.color = cssv('--text-2');
  Chart.defaults.borderColor = cssv('--line');
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15,35,64,.94)';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 8;
  Chart.defaults.plugins.legend.labels.boxWidth = 12;
  Chart.defaults.plugins.legend.labels.boxHeight = 12;
  Chart.defaults.plugins.legend.labels.useBorderRadius = true;
  Chart.defaults.plugins.legend.labels.borderRadius = 3;
  Chart.defaults.animation.duration = 450;
}
// Plugin vẽ nhãn số trên cột / tổng cột chồng
var valuePlugin = {
  id: 'vlabels',
  afterDatasetsDraw: function (chart, args, opts) {
    if (!opts || !opts.on) return;
    var ctx = chart.ctx, horiz = chart.options.indexAxis === 'y';
    var metas = chart.data.datasets.map(function (d, i) { return chart.getDatasetMeta(i); }).filter(function (m) { return !m.hidden && m.type === 'bar'; });
    if (!metas.length) return;
    var labels = chart.data.labels, xs = chart.scales.x, ys = chart.scales.y;
    ctx.save(); ctx.font = '700 11px "Segoe UI",Arial'; ctx.fillStyle = cssv('--text');
    labels.forEach(function (l, i) {
      var tot = 0; metas.forEach(function (m) { var v = chart.data.datasets[m.index].data[i]; if (typeof v === 'number' && m.stack !== 'line') tot += v; });
      if (!tot) return;
      var txt = opts.fmt ? opts.fmt(tot, i) : fmtN(tot);
      if (opts.stacked === false) {
        metas.forEach(function (m) {
          var v = chart.data.datasets[m.index].data[i], el = m.data[i]; if (!v || !el) return;
          var t = opts.fmt ? opts.fmt(v) : fmtN(v);
          if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(t, el.x + 4, el.y); }
          else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(t, el.x, el.y - 3); }
        });
        return;
      }
      if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, xs.getPixelForValue(tot) + 5, ys.getPixelForValue(i)); }
      else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, xs.getPixelForValue(i), ys.getPixelForValue(tot) - 3); }
    });
    ctx.restore();
  }
};
var centerPlugin = {
  id: 'center',
  afterDraw: function (chart, args, opts) {
    if (!opts || !opts.text) return;
    var a = chart.chartArea, ctx = chart.ctx, x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
    ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    ctx.fillStyle = opts.color || cssv('--text'); ctx.font = '800 ' + (opts.size || 26) + 'px "Segoe UI",Arial'; ctx.fillText(opts.text, x, y - 7);
    ctx.font = '700 10.5px "Segoe UI",Arial'; ctx.globalAlpha = .75; ctx.fillText(opts.sub || '', x, y + 15); ctx.restore();
  }
};
var arcLabelPlugin = {
  id: 'arcl',
  afterDatasetsDraw: function (chart, args, opts) {
    if (!opts || !opts.on) return;
    var ds = chart.data.datasets[0], meta = chart.getDatasetMeta(0), tot = ds.data.reduce(function (a, b) { return a + b; }, 0);
    var ctx = chart.ctx; ctx.save(); ctx.font = '700 11px "Segoe UI",Arial'; ctx.fillStyle = '#fff'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    meta.data.forEach(function (el, i) { var v = ds.data[i]; if (!v || v / tot < .06 || meta.hidden || !chart.getDataVisibility(i)) return; var p = el.tooltipPosition(); ctx.fillText(Math.round(v / tot * 100) + '%', p.x, p.y); });
    ctx.restore();
  }
};
function mkChart(id, cfg) {
  var cv = document.getElementById(id); if (!cv) return null;
  if (CH[id]) { var ch = CH[id]; if (ch.config.type === cfg.type) { ch.data.labels = cfg.data.labels; ch.data.datasets = cfg.data.datasets; ch.options = cfg.options; ch.update(); return ch; } ch.destroy(); }
  cfg.plugins = [valuePlugin, centerPlugin, arcLabelPlugin];
  CH[id] = new Chart(cv, cfg); return CH[id];
}
function clickTo(fk, labelsRef, mapFn) {
  return function (evt, els, chart) {
    if (!els.length) return;
    var v = (labelsRef || chart.data.labels)[els[0].index];
    if (mapFn) v = mapFn(v, els[0]);
    if (v == null) return;
    if (typeof fk === 'function') fk(v, els[0]); else toggleFacet(fk, v, true);
  };
}
function donut(id, labels, data, colors, centerText, centerSub, fk, onDark) {
  mkChart(id, {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: data, backgroundColor: colors, borderColor: onDark ? 'rgba(255,255,255,.9)' : cssv('--card'), borderWidth: 2, hoverOffset: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '68%',
      plugins: { legend: { display: false }, center: { text: centerText, sub: centerSub, color: onDark ? '#fff' : null }, arcl: { on: !onDark } },
      onClick: fk ? clickTo(fk) : null, onHover: hoverCursor
    }
  });
}
function hoverCursor(e, els) { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
function pieCard(id, entries, colorOf, fk) {
  var cv = document.getElementById(id), w = cv ? cv.parentNode.clientWidth : 600, side = w >= 560;
  var tot = entries.reduce(function (a, e) { return a + e[1]; }, 0);
  mkChart(id, {
    type: 'doughnut',
    data: { labels: entries.map(function (e) { return e[0]; }), datasets: [{ data: entries.map(function (e) { return e[1]; }), backgroundColor: entries.map(function (e, i) { return colorOf(e[0], i); }), borderColor: cssv('--card'), borderWidth: 2, hoverOffset: 6 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '58%',
      plugins: {
        legend: { position: side ? 'right' : 'bottom', labels: { generateLabels: function (ch) { var d = ch.data; return d.labels.map(function (l, i) { return { text: shortName(l, side ? 34 : 26) + '  ' + d.datasets[0].data[i] + ' (' + pct(d.datasets[0].data[i], tot, 0) + ')', fillStyle: d.datasets[0].backgroundColor[i], strokeStyle: d.datasets[0].backgroundColor[i], hidden: !ch.getDataVisibility(i), index: i, fontColor: cssv('--text-2') }; }); } } },
        center: { text: fmtN(tot), sub: 'PHIẾU', size: 20 }, arcl: { on: true },
        tooltip: { callbacks: { label: function (c) { return ' ' + c.label + ': ' + c.raw + ' (' + pct(c.raw, tot) + ')'; } } }
      },
      onClick: clickTo(fk), onHover: hoverCursor
    }
  });
}
function barOpts(horiz, stacked, extra) {
  var o = {
    responsive: true, maintainAspectRatio: false, indexAxis: horiz ? 'y' : 'x',
    scales: {
      x: { stacked: stacked, grid: { display: !!horiz, color: cssv('--line') }, ticks: { autoSkip: false, maxRotation: horiz ? 0 : 40, minRotation: 0 } },
      y: { stacked: stacked, grid: { display: !horiz, color: cssv('--line') }, ticks: { autoSkip: false }, beginAtZero: true }
    },
    plugins: { legend: { position: 'bottom' }, vlabels: { on: true }, tooltip: { mode: 'index', intersect: false } },
    interaction: { mode: 'index', intersect: false, axis: horiz ? 'y' : 'x' },
    onHover: hoverCursor, layout: { padding: horiz ? { right: 34 } : { top: 18 } }
  };
  return Object.assign(o, extra || {});
}
function stackDatasets(rows, catFn, catKeys, serFn, serKeys, colorOf) {
  var m = {}; rows.forEach(function (r) { var c = catFn(r), s = serFn(r); (m[c] = m[c] || {})[s] = (m[c][s] || 0) + 1; });
  return serKeys.map(function (s, i) {
    return { label: s, data: catKeys.map(function (c) { return (m[c] && m[c][s]) || 0; }), backgroundColor: colorOf(s, i), borderRadius: 4, borderSkipped: false, maxBarThickness: 34 };
  }).filter(function (d) { return d.data.some(function (v) { return v; }); });
}
function shortName(s, n) { s = String(s); n = n || 26; return s.length > n ? s.slice(0, n - 1) + '…' : s; }

function renderCharts() {
  var R = S.filtered;
  // Hãng
  var hk = sortedEntries(countBy(R, function (r) { return r.hang; })).map(function (e) { return e[0]; });
  var hSer = S.hangMode === 'flow' ? ['Trực tiếp', 'Exwork'] : ['Hoàn thành', 'Đang xử lý'];
  var hds = stackDatasets(R, function (r) { return r.hang; }, hk, function (r) { return S.hangMode === 'flow' ? r.flow : r.st; }, hSer, function (s) { return (S.hangMode === 'flow' ? FLOW_COLOR : ST_COLOR)[s] || '#94A3B8'; });
  $('#cHang').parentNode.style.height = Math.max(260, hk.length * 24 + 70) + 'px';
  mkChart('cHang', { type: 'bar', data: { labels: hk, datasets: hds }, options: barOpts(true, true, { onClick: clickTo('hang') }) });
  // Sale
  var sk = sortedEntries(countBy(R, function (r) { return r.sale; })).map(function (e) { return e[0]; });
  mkChart('cSale', { type: 'bar', data: { labels: sk, datasets: stackDatasets(R, function (r) { return r.sale; }, sk, function (r) { return r.st; }, ['Hoàn thành', 'Đang xử lý'], function (s) { return ST_COLOR[s]; }) }, options: barOpts(false, true, { onClick: clickTo('sale') }) });
  CH.cSale.options.scales.x.ticks.callback = function (v, i) { return shortName(this.getLabelForValue(v), 18); }; CH.cSale.update('none');
  // Tháng × top hãng
  var mk = Array.from(new Set(R.map(function (r) { return r.month; }))).sort();
  var top = hk.slice(0, 6), mser = top.concat(hk.length > 6 ? ['Khác'] : []);
  var mds = stackDatasets(R, function (r) { return r.month; }, mk, function (r) { return top.indexOf(r.hang) >= 0 ? r.hang : 'Khác'; }, mser, function (s) { return s === 'Khác' ? '#CBD5E1' : catColor(s); });
  mkChart('cMonth', { type: 'bar', data: { labels: mk.map(ymLabel), datasets: mds }, options: barOpts(false, true, { onClick: clickTo('month', mk) }) });
  // Thị phần
  var he = sortedEntries(countBy(R, function (r) { return r.hang; }));
  var share = he.slice(0, 8); if (he.length > 8) share.push(['Khác', he.slice(8).reduce(function (a, e) { return a + e[1]; }, 0)]);
  pieCard('cShare', share, function (k) { return k === 'Khác' ? '#CBD5E1' : catColor(k); }, function (v) { if (v !== 'Khác') toggleFacet('hang', v, true); });
  // PM
  var FL = flagLabels(), fk = ['m3', 'm2', 'm1', 'ok', 'done'];
  var pk = sortedEntries(countBy(R, function (r) { return r.pm; })).map(function (e) { return e[0]; });
  mkChart('cPM', { type: 'bar', data: { labels: pk, datasets: stackDatasets(R, function (r) { return r.pm; }, pk, function (r) { return r.flagLabel; }, fk.map(function (k) { return FL[k]; }), function (s, i) { return FLAG_META[fk[i]].c; }) }, options: barOpts(false, true, { onClick: clickTo('pm') }) });
  CH.cPM.options.scales.x.ticks.callback = function (v) { return shortName(this.getLabelForValue(v), 18); }; CH.cPM.update('none');
  // Bước
  var stp = Array.from(countBy(R, function (r) { return r.step; }).entries()).sort(function (a, b) { return stepIdx(a[0]) - stepIdx(b[0]); });
  mkChart('cStep', {
    type: 'bar', data: { labels: stp.map(function (e) { return e[0]; }), datasets: [{ label: 'Số phiếu', data: stp.map(function (e) { return e[1]; }), backgroundColor: stp.map(function (e, i) { return e[0] === 'Kết thúc' ? '#16A34A' : ['#94A3B8', '#64748B', '#6366F1', '#8B5CF6', '#0EA5E9', '#F59E0B', '#F97316', '#EF4444'][i % 8]; }), borderRadius: 5, maxBarThickness: 26 }] },
    options: barOpts(true, false, { onClick: clickTo('step'), plugins: { legend: { display: false }, vlabels: { on: true }, tooltip: {} } })
  });
  // Thời gian xử lý theo hãng
  var tgm = {}; R.forEach(function (r) { if (r.st === 'Hoàn thành' && r.tg != null) (tgm[r.hang] = tgm[r.hang] || []).push(r.tg); });
  var tk = Object.keys(tgm).sort(function (a, b) { return tgm[b].length - tgm[a].length; }).slice(0, 14).sort(function (a, b) { return mean(tgm[b]) - mean(tgm[a]); });
  mkChart('cTG', {
    type: 'bar', data: {
      labels: tk, datasets: [
        { label: 'Trung bình (ngày)', data: tk.map(function (k) { return Math.round(mean(tgm[k]) * 10) / 10; }), backgroundColor: '#0959A2', borderRadius: 4, maxBarThickness: 22 },
        { label: 'Trung vị (ngày)', data: tk.map(function (k) { return Math.round(median(tgm[k]) * 10) / 10; }), backgroundColor: '#00A8E8', borderRadius: 4, maxBarThickness: 22 }]
    },
    options: barOpts(false, false, { onClick: clickTo('hang'), plugins: { legend: { position: 'bottom' }, vlabels: { on: true, stacked: false, fmt: function (v) { return fmtN(v, 1); } }, tooltip: { mode: 'index', intersect: false, callbacks: { afterTitle: function (it) { return 'Số phiếu HT: ' + tgm[it[0].label].length; } } } } })
  });
  renderCal();
  // Nhóm KH, Luồng, HĐ
  pieCard('cNhom', sortedEntries(countBy(R, function (r) { return r.nhom; })), function (k, i) { return PALETTE[(i + 1) % PALETTE.length]; }, 'nhom');
  pieCard('cFlow', sortedEntries(countBy(R, function (r) { return r.flow; })), function (k) { return FLOW_COLOR[k] || '#94A3B8'; }, 'flow');
  pieCard('cHD', sortedEntries(countBy(R, function (r) { return r.hd; })), function (k) { return HD_COLOR[k] || '#94A3B8'; }, 'hd');
  // Đại lý & KH
  var nTop = +$('#dlTop').value || 15;
  var dle = sortedEntries(countBy(R.filter(function (r) { return r.dl !== EMPTY; }), function (r) { return r.dl; })).slice(0, nTop);
  $('#cDL').parentNode.style.height = Math.max(300, dle.length * 22 + 60) + 'px';
  mkChart('cDL', { type: 'bar', data: { labels: dle.map(function (e) { return shortName(e[0], 34); }), datasets: stackDatasets(R, function (r) { return r.dl; }, dle.map(function (e) { return e[0]; }), function (r) { return r.st; }, ['Hoàn thành', 'Đang xử lý'], function (s) { return s === 'Hoàn thành' ? '#0959A2' : '#F06723'; }) }, options: barOpts(true, true, { onClick: clickTo('dl', dle.map(function (e) { return e[0]; })) }) });
  var khe = sortedEntries(countBy(R.filter(function (r) { return r.kh !== EMPTY; }), function (r) { return r.kh; })).slice(0, 15);
  mkChart('cKH', { type: 'bar', data: { labels: khe.map(function (e) { return shortName(e[0], 34); }), datasets: [{ label: 'Số phiếu', data: khe.map(function (e) { return e[1]; }), backgroundColor: '#14B8A6', borderRadius: 4, maxBarThickness: 20 }] }, options: barOpts(true, false, { onClick: clickTo('kh', khe.map(function (e) { return e[0]; })), plugins: { legend: { display: false }, vlabels: { on: true }, tooltip: { callbacks: { title: function (it) { return khe[it[0].dataIndex][0]; } } } } }) });
}
function renderCal() {
  var R = S.filtered, key = { giao: 'dGiao', le: 'dLe', ls: 'dLs' }[S.calMode];
  var m = {}; R.forEach(function (r) { var t = r[key]; if (t == null) return; var ym = C.ymOf(t); m[ym] = m[ym] || {}; var s = S.calMode === 'giao' ? r.st : r.hang; m[ym][s] = (m[ym][s] || 0) + 1; });
  var ks = Object.keys(m).sort();
  if (ks.length > 24) { var refYm = C.ymOf(S.ref), i0 = Math.max(0, ks.findIndex(function (k) { return k >= refYm; }) - 6); ks = ks.slice(i0 < 0 ? 0 : i0, (i0 < 0 ? 0 : i0) + 24); }
  var ser;
  if (S.calMode === 'giao') ser = ['Hoàn thành', 'Đang xử lý'];
  else { var hc = {}; ks.forEach(function (k) { for (var h in m[k]) hc[h] = (hc[h] || 0) + m[k][h]; }); ser = Object.keys(hc).sort(function (a, b) { return hc[b] - hc[a]; }).slice(0, 6); var hasOther = Object.keys(hc).length > 6; if (hasOther) ser.push('Khác'); }
  var ds = ser.map(function (s) {
    return { label: s, data: ks.map(function (k) { if (s === 'Khác') { var t = 0; for (var h in m[k]) if (ser.indexOf(h) < 0) t += m[k][h]; return t; } return m[k][s] || 0; }), backgroundColor: S.calMode === 'giao' ? ST_COLOR[s] : s === 'Khác' ? '#CBD5E1' : catColor(s), borderRadius: 4, maxBarThickness: 30 };
  });
  var refYm2 = C.ymOf(S.ref);
  mkChart('cCal', {
    type: 'bar', data: { labels: ks.map(function (k) { return ymLabel(k) + (k === refYm2 ? ' ◆' : ''); }), datasets: ds },
    options: barOpts(false, true, {
      onClick: clickTo(function (v, el) {
        var ym = ks[el.index].split('-'); var y = +ym[0], mo = +ym[1] - 1;
        setDate(key, { from: monthStart(y, mo), to: monthEnd(y, mo), label: 'T' + (mo + 1) + '/' + y });
      })
    })
  });
}

/* ----------------------------- bảng thống kê ----------------------------- */
function renderStatTables() {
  var R = S.filtered;
  // Tuân thủ theo Sale
  var by = {};
  R.forEach(function (r) {
    var o = by[r.sale] = by[r.sale] || { n: 0, done: 0, ky: 0, bg: 0, ret: 0, no: 0, bbC: 0, bbK: 0, bbE: 0 };
    o.n++; if (r.st !== 'Hoàn thành') return; o.done++;
    if (r.hd === 'Đã ký') o.ky++; else if (r.hd === 'Đã ký và bàn giao cho kế toán') o.bg++; else if (r.hd === 'KH đã ký nhưng chưa trả Hợp đồng') o.ret++; else if (r.hd === 'Chưa ký') o.no++;
    if (r.bb === 'Có') o.bbC++; else if (r.bb === 'Không') o.bbK++; else o.bbE++;
  });
  var keys = Object.keys(by).sort(function (a, b) { return by[b].done - by[a].done || by[b].n - by[a].n; });
  var T = { n: 0, done: 0, ky: 0, bg: 0, ret: 0, no: 0, bbC: 0, bbK: 0, bbE: 0 };
  keys.forEach(function (k) { for (var f in T) T[f] += by[k][f]; });
  function rate(o) { var r = o.done ? (o.ky + o.bg) / o.done : null; var c = r == null ? '#CBD5E1' : r >= .9 ? '#16A34A' : r >= .7 ? '#F59E0B' : '#DC2626'; return (r == null ? '–' : fmtN(r * 100, 0) + '%') + '<span class="pbar"><i style="width:' + ((r || 0) * 100).toFixed(0) + '%;background:' + c + '"></i></span>'; }
  var h = '<table class="mtable"><thead><tr><th>Sale</th><th class="n">Tổng</th><th class="n">HT</th><th class="n">Đã ký</th><th class="n">Ký & BG KT</th><th class="n">KH chưa trả</th><th class="n">Chưa ký</th><th class="n">Tuân thủ HĐ</th><th class="n">BB Có</th><th class="n">BB Không</th><th class="n">BB trống</th></tr></thead><tbody>';
  keys.forEach(function (k) { var o = by[k]; h += '<tr class="click" data-f="sale" data-v="' + esc(k) + '"><td>' + esc(k) + '</td><td class="n">' + o.n + '</td><td class="n">' + o.done + '</td><td class="n">' + o.ky + '</td><td class="n">' + o.bg + '</td><td class="n">' + (o.ret ? '<b style="color:#B45309">' + o.ret + '</b>' : 0) + '</td><td class="n">' + (o.no ? '<b style="color:#DC2626">' + o.no + '</b>' : 0) + '</td><td class="n">' + rate(o) + '</td><td class="n">' + o.bbC + '</td><td class="n">' + o.bbK + '</td><td class="n">' + o.bbE + '</td></tr>'; });
  h += '</tbody><tfoot><tr><td>Tổng</td><td class="n">' + T.n + '</td><td class="n">' + T.done + '</td><td class="n">' + T.ky + '</td><td class="n">' + T.bg + '</td><td class="n">' + T.ret + '</td><td class="n">' + T.no + '</td><td class="n">' + rate(T) + '</td><td class="n">' + T.bbC + '</td><td class="n">' + T.bbK + '</td><td class="n">' + T.bbE + '</td></tr></tfoot></table>';
  $('#tCompliance').innerHTML = keys.length ? h : '<div class="empty-state">Không có dữ liệu</div>';
  // Theo hãng
  var hs = hangStats(R);
  var hh = '<table class="mtable"><thead><tr><th>Hãng</th><th class="n">Tổng</th><th class="n">Tỷ trọng</th><th class="n">Đang XL</th><th class="n">HT</th><th class="n">TG TB</th><th class="n">TG trung vị</th><th class="n">Cần nhắc</th><th class="n">Trễ giao</th><th class="n">HĐ chưa ký</th><th class="n">Exw</th><th class="n">Số KH</th></tr></thead><tbody>';
  hs.forEach(function (o) { hh += '<tr class="click" data-f="hang" data-v="' + esc(o.k) + '"><td><i style="display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;background:' + catColor(o.k) + '"></i>' + esc(o.k) + '</td><td class="n"><b>' + o.n + '</b></td><td class="n">' + pct(o.n, R.length) + '<span class="pbar"><i style="width:' + (o.n / (hs[0] ? hs[0].n : 1) * 100).toFixed(0) + '%;background:' + catColor(o.k) + '"></i></span></td><td class="n">' + o.ip + '</td><td class="n">' + o.done + '</td><td class="n">' + fmtN(o.avg, 1) + '</td><td class="n">' + fmtN(o.med, 1) + '</td><td class="n">' + (o.flag ? '<b style="color:#F97316">' + o.flag + '</b>' : 0) + '</td><td class="n">' + (o.late ? '<b style="color:#DC2626">' + o.late + '</b>' : 0) + '</td><td class="n">' + o.hdNo + '</td><td class="n">' + o.exw + '</td><td class="n">' + o.kh + '</td></tr>'; });
  hh += '</tbody></table>';
  $('#tHang').innerHTML = hs.length ? hh : '<div class="empty-state">Không có dữ liệu</div>';
  // Heatmap
  var months = Array.from(new Set(R.map(function (r) { return r.month; }))).sort();
  var mx = 0, cell = {}; R.forEach(function (r) { var k = r.hang + '|' + r.month; cell[k] = (cell[k] || 0) + 1; if (cell[k] > mx) mx = cell[k]; });
  var ht = '<table class="mtable heat"><thead><tr><th>Hãng \\ Tháng</th>' + months.map(function (m) { return '<th class="n">' + ymLabel(m) + '</th>'; }).join('') + '<th class="n">Tổng</th></tr></thead><tbody>';
  hs.forEach(function (o) {
    ht += '<tr><td>' + esc(o.k) + '</td>' + months.map(function (m) { var v = cell[o.k + '|' + m] || 0; var a = v ? .12 + .78 * v / mx : 0; return '<td class="h click" data-hm="' + esc(o.k) + '|' + esc(m) + '" style="cursor:' + (v ? 'pointer' : 'default') + ';background:rgba(9,89,162,' + a.toFixed(2) + ');color:' + (a > .5 ? '#fff' : 'inherit') + '">' + (v || '') + '</td>'; }).join('') + '<td class="n"><b>' + o.n + '</b></td></tr>';
  });
  var colT = months.map(function (m) { return R.filter(function (r) { return r.month === m; }).length; });
  ht += '</tbody><tfoot><tr><td>Tổng</td>' + colT.map(function (v) { return '<td class="n">' + v + '</td>'; }).join('') + '<td class="n">' + R.length + '</td></tr></tfoot></table>';
  $('#tHeat').innerHTML = hs.length ? ht : '<div class="empty-state">Không có dữ liệu</div>';
}
function hangStats(R) {
  var by = {};
  R.forEach(function (r) {
    var o = by[r.hang] = by[r.hang] || { k: r.hang, n: 0, ip: 0, done: 0, tg: [], flag: 0, late: 0, hdNo: 0, exw: 0, khs: new Set(), dls: new Set(), lic: 0 };
    o.n++; if (r.st === 'Hoàn thành') { o.done++; if (r.tg != null) o.tg.push(r.tg); } else o.ip++;
    if (r.flag === 'm1' || r.flag === 'm2' || r.flag === 'm3') o.flag++; if (r.giao === 'late') o.late++; if (r.hd === 'Chưa ký') o.hdNo++; if (r.flow === 'Exwork') o.exw++; if (r.lic === 'soon') o.lic++;
    o.khs.add(r.kh); if (r.dl !== EMPTY) o.dls.add(r.dl);
  });
  return Object.keys(by).map(function (k) { var o = by[k]; o.avg = mean(o.tg); o.med = median(o.tg); o.kh = o.khs.size; o.dl = o.dls.size; return o; }).sort(function (a, b) { return b.n - a.n; });
}

/* ----------------------------- bảng chi tiết ----------------------------- */
function renderTable() {
  var cols = visCols.map(function (k) { return COL_BY[k]; }).filter(Boolean);
  var R = S.filtered, ps = S.pageSize, pages = ps ? Math.max(1, Math.ceil(R.length / ps)) : 1;
  if (S.page > pages) S.page = pages;
  var start = ps ? (S.page - 1) * ps : 0, rows = ps ? R.slice(start, start + ps) : R;
  $('#dt').className = 'dt' + (S.density === 'compact' ? ' compact' : '');
  $('#dt thead').innerHTML = '<tr>' + cols.map(function (c) { var on = S.sort.k === c.k; return '<th data-sort="' + c.k + '" class="' + (on ? 'sorted ' : '') + (c.n ? 'n' : '') + '">' + esc(c.label) + '<span class="so">' + (on ? (S.sort.d > 0 ? '▲' : '▼') : '⇅') + '</span></th>'; }).join('') + '</tr>';
  $('#dt tbody').innerHTML = rows.length ? rows.map(function (r) {
    return '<tr data-i="' + r._i + '">' + cols.map(function (c) { var t = cellText(r, c); return '<td class="' + (c.n ? 'n' : '') + (c.w ? ' wrap' : '') + '" title="' + esc(t == null ? '' : t) + '">' + cellHtml(r, c) + '</td>'; }).join('') + '</tr>';
  }).join('') : '<tr><td colspan="' + cols.length + '"><div class="empty-state">Không có phiếu nào khớp bộ lọc. <a href="#" data-rm="*">Xóa bộ lọc</a></div></td></tr>';
  $('#tblCount').innerHTML = 'Hiển thị <b>' + fmtN(rows.length ? start + 1 : 0) + '–' + fmtN(start + rows.length) + '</b> trong <b>' + fmtN(R.length) + '</b> phiếu · ' + cols.length + '/' + COLS.length + ' cột';
  var p = '';
  if (ps && pages > 1) {
    p += '<button data-pg="' + (S.page - 1) + '"' + (S.page === 1 ? ' disabled' : '') + '>‹</button>';
    var arr = []; for (var i = 1; i <= pages; i++) if (i === 1 || i === pages || Math.abs(i - S.page) <= 2) arr.push(i);
    var last = 0; arr.forEach(function (i) { if (i - last > 1) p += '<span>…</span>'; p += '<button data-pg="' + i + '" class="' + (i === S.page ? 'on' : '') + '">' + i + '</button>'; last = i; });
    p += '<button data-pg="' + (S.page + 1) + '"' + (S.page === pages ? ' disabled' : '') + '>›</button>';
  }
  $('#pager').innerHTML = '<span style="margin-right:auto">Trang ' + S.page + '/' + pages + '</span>' + p;
}

/* ----------------------------- popovers ----------------------------- */
var openPop = null; // {el, kind, key}
function closePop() { if (openPop) { var p = $('.pop', openPop.el); if (p) p.remove(); openPop = null; } }
function placePop(pop, host) {
  var r = host.getBoundingClientRect();
  if (r.left + 400 > window.innerWidth) pop.classList.add('right');
}
function refreshOpenPop() {
  if (!openPop) return;
  if (openPop.kind === 'facet') fillFacetList(openPop);
  else if (openPop.kind === 'date') fillDatePop(openPop);
}
function openFacetPop(host, k) {
  if (openPop && openPop.el === host) { closePop(); return; }
  closePop();
  var f = FACET_BY[k], pop = document.createElement('div'); pop.className = 'pop';
  pop.innerHTML = '<div class="ps"><svg class="i"><use href="#i-search"/></svg><input type="text" placeholder="Tìm trong ' + esc(f.label.toLowerCase()) + '…"></div><div class="list"></div>' +
    '<div class="pf"><button class="btn sm ghost" data-a="all">Chọn tất cả (kết quả)</button><button class="btn sm ghost" data-a="none">Bỏ chọn</button><button class="btn sm primary" data-a="close">Xong</button></div>';
  host.appendChild(pop); placePop(pop, host);
  openPop = { el: host, kind: 'facet', key: k, q: '' };
  var inp = $('input', pop);
  inp.addEventListener('input', function () { openPop.q = inp.value; fillFacetList(openPop); });
  inp.addEventListener('keydown', function (e) { if (e.key === 'Enter') { var first = $('label.opt input', pop); if (first) { first.click(); } } });
  pop.addEventListener('click', function (e) {
    var a = e.target.closest('[data-a]'); if (!a) return;
    if (a.dataset.a === 'close') { closePop(); return; }
    var s = F.facets[k];
    if (a.dataset.a === 'none') s.clear();
    else $$('label.opt', pop).forEach(function (l) { s.add(l.dataset.v); });
    S.page = 1; update();
  });
  pop.addEventListener('change', function (e) {
    var l = e.target.closest('label.opt'); if (!l) return;
    var s = F.facets[k]; if (e.target.checked) s.add(l.dataset.v); else s.delete(l.dataset.v);
    S.page = 1; update();
  });
  fillFacetList(openPop);
  setTimeout(function () { inp.focus(); }, 20);
}
function fillFacetList(op) {
  var k = op.key, f = FACET_BY[k], pop = $('.pop', op.el); if (!pop) return;
  var all = countBy(S.rows, function (r) { return r[k]; });
  var cur = countBy(S.rows.filter(function (r) { return passes(r, k); }), function (r) { return r[k]; });
  var items = Array.from(all.keys());
  if (f.fixed) { var ord = k === 'flagLabel' ? ['m3', 'm2', 'm1', 'ok', 'done'].map(function (x) { return flagLabels()[x]; }) : null; if (ord) items.sort(function (a, b) { return ord.indexOf(a) - ord.indexOf(b); }); }
  else if (f.order) items.sort(function (a, b) { return stepIdx(a) - stepIdx(b); });
  else if (f.sortKey) items.sort().reverse();
  else items.sort(function (a, b) { return (cur.get(b) || 0) - (cur.get(a) || 0) || (all.get(b) - all.get(a)) || String(a).localeCompare(String(b), 'vi'); });
  var q = nk(op.q || ''), sel = F.facets[k];
  if (q) items = items.filter(function (v) { return nk(v).indexOf(q) >= 0; });
  // các mục đang chọn lên đầu
  items.sort(function (a, b) { return (sel.has(b) ? 1 : 0) - (sel.has(a) ? 1 : 0); });
  var list = $('.list', pop);
  if (!items.length) { list.innerHTML = '<div class="empty">Không tìm thấy “' + esc(op.q) + '”</div>'; return; }
  list.innerHTML = items.slice(0, 400).map(function (v) {
    var c = cur.get(v) || 0, lbl = k === 'month' ? ymLabel(v) : v;
    return '<label class="opt' + (c || sel.has(v) ? '' : ' dim') + '" data-v="' + esc(v) + '"><input type="checkbox"' + (sel.has(v) ? ' checked' : '') + '><span class="tx" title="' + esc(lbl) + '">' + hl(lbl, op.q) + '</span><span class="ct">' + c + '</span></label>';
  }).join('') + (items.length > 400 ? '<div class="empty">… còn ' + (items.length - 400) + ' mục, hãy gõ để thu hẹp</div>' : '');
}
function hl(text, q) {
  text = String(text); if (!q) return esc(text);
  var nq = nk(q); if (!nq) return esc(text);
  // ánh xạ vị trí ký tự không dấu → gốc
  var norm = '', map = [];
  for (var i = 0; i < text.length; i++) { var c = nk(text[i]) || (text[i] === ' ' ? ' ' : ''); for (var j = 0; j < c.length; j++) { norm += c[j]; map.push(i); } }
  var out = '', last = 0, idx, from = 0, toks = nq.split(' ').filter(Boolean), ranges = [];
  toks.forEach(function (t) { from = 0; while ((idx = norm.indexOf(t, from)) >= 0) { ranges.push([map[idx], map[idx + t.length - 1] + 1]); from = idx + t.length; } });
  ranges.sort(function (a, b) { return a[0] - b[0]; });
  ranges.forEach(function (r) { if (r[0] < last) return; out += esc(text.slice(last, r[0])) + '<mark>' + esc(text.slice(r[0], r[1])) + '</mark>'; last = r[1]; });
  return out + esc(text.slice(last));
}

/* ---- popover ngày ---- */
function openDatePop(host, k) {
  if (openPop && openPop.el === host) { closePop(); return; }
  closePop();
  var pop = document.createElement('div'); pop.className = 'pop'; pop.style.minWidth = '360px';
  host.appendChild(pop); placePop(pop, host);
  var d = F.dates[k], ys = years(k);
  var refY = new Date(S.ref).getUTCFullYear();
  openPop = { el: host, kind: 'date', key: k, year: d && d.from != null ? new Date(d.from).getUTCFullYear() : (ys.indexOf(refY) >= 0 ? refY : ys[ys.length - 1] || refY), anchor: null };
  pop.addEventListener('click', function (e) { onDatePopClick(e, openPop); });
  pop.addEventListener('change', function (e) {
    var t = e.target;
    if (t.matches('input[type=date]')) {
      var f = parseIso($('input[data-r=from]', pop).value), to = parseIso($('input[data-r=to]', pop).value);
      if (f == null && to == null) setDate(k, null); else setDate(k, { from: f, to: to });
    } else if (t.matches('input[data-a=empty]')) setDate(k, t.checked ? { empty: true } : null);
  });
  fillDatePop(openPop);
}
function years(k) { var s = new Set(); S.rows.forEach(function (r) { if (r[k] != null) s.add(new Date(r[k]).getUTCFullYear()); }); return Array.from(s).sort(); }
function presetsFor(k) {
  var R = S.refDay, d = new Date(R), y = d.getUTCFullYear(), m = d.getUTCMonth(), q = Math.floor(m / 3);
  var P = [
    ['Tháng này', monthStart(y, m), monthEnd(y, m)], ['Tháng trước', monthStart(y, m - 1), monthEnd(y, m - 1)],
    ['7 ngày qua', R - 6 * DAY, R], ['30 ngày qua', R - 29 * DAY, R], ['90 ngày qua', R - 89 * DAY, R],
    ['Quý này', monthStart(y, q * 3), monthEnd(y, q * 3 + 2)], ['Quý trước', monthStart(y, q * 3 - 3), monthEnd(y, q * 3 - 1)],
    ['Năm nay', Date.UTC(y, 0, 1), Date.UTC(y, 11, 31)], ['Năm trước', Date.UTC(y - 1, 0, 1), Date.UTC(y - 1, 11, 31)]
  ];
  if (!DATE_BY[k].past) P = P.concat([
    ['Đã qua (trước đối chiếu)', null, R - DAY], ['7 ngày tới', R, R + 7 * DAY], ['30 ngày tới', R, R + 30 * DAY], ['90 ngày tới', R, R + 90 * DAY],
    ['6 tháng tới', R, Date.UTC(y, m + 6, d.getUTCDate())], ['Tháng sau', monthStart(y, m + 1), monthEnd(y, m + 1)], ['Từ hôm nay trở đi', R, null]
  ]);
  return P;
}
function fillDatePop(op) {
  var k = op.key, pop = $('.pop', op.el); if (!pop) return;
  var d = F.dates[k] || {}, y = op.year;
  var cnt = {}; S.rows.forEach(function (r) { if (r[k] != null) { var t = new Date(r[k]); if (t.getUTCFullYear() === y) cnt[t.getUTCMonth()] = (cnt[t.getUTCMonth()] || 0) + 1; } });
  var empties = S.rows.filter(function (r) { return r[k] == null; }).length;
  var h = '<h5>Chọn nhanh</h5><div class="presets">' + presetsFor(k).map(function (p, i) { return '<button data-p="' + i + '" class="' + (d.label === p[0] ? 'on' : '') + '">' + p[0] + '</button>'; }).join('') + '</div>';
  h += '<h5 style="margin-top:12px">Theo tháng / quý (bấm 2 tháng để chọn khoảng)</h5><div class="yr"><button class="btn sm icon ghost" data-y="-1">' + icon('left') + '</button><b>Năm ' + y + '</b><button class="btn sm icon ghost" data-y="1">' + icon('right') + '</button></div><div class="mgrid">';
  for (var m = 0; m < 12; m++) {
    var ms = monthStart(y, m), me = monthEnd(y, m);
    var inR = d.from != null && d.to != null && ms >= dayStart(d.from) && me <= d.to;
    var exact = d.from === ms && d.to === me;
    h += '<button data-m="' + m + '" class="' + (cnt[m] ? 'has ' : '') + (exact ? 'on' : inR ? 'in' : '') + '" title="' + (cnt[m] || 0) + ' phiếu">T' + (m + 1) + '</button>';
  }
  h += '</div><div class="presets" style="margin-top:6px">' + [0, 1, 2, 3].map(function (q) { return '<button data-q="' + q + '">Q' + (q + 1) + '/' + y + '</button>'; }).join('') + '<button data-q="y">Cả năm ' + y + '</button></div>';
  h += '<h5 style="margin-top:12px">Khoảng tùy chọn (lịch)</h5><div class="drange"><label>Từ ngày<input type="date" data-r="from" value="' + (d.from != null ? isoD(d.from) : '') + '"></label><label>Đến ngày<input type="date" data-r="to" value="' + (d.to != null ? isoD(d.to) : '') + '"></label></div>';
  h += '<label class="chk"><input type="checkbox" data-a="empty"' + (d.empty ? ' checked' : '') + '> Chỉ phiếu chưa có ngày (' + empties + ')</label>';
  h += '<div class="pf"><button class="btn sm ghost" data-a="clear">Xóa lọc ngày</button><button class="btn sm primary" data-a="close">Xong</button></div>';
  pop.innerHTML = h;
}
function onDatePopClick(e, op) {
  var k = op.key, b = e.target.closest('button'); if (!b) return;
  if (b.dataset.y) { op.year += +b.dataset.y; fillDatePop(op); return; }
  if (b.dataset.p != null) { var p = presetsFor(k)[+b.dataset.p]; setDate(k, { from: p[1], to: p[2], label: p[0] }); return; }
  if (b.dataset.m != null) {
    var m = +b.dataset.m, y = op.year, ms = monthStart(y, m), me = monthEnd(y, m);
    if (op.anchor && op.anchor[0] !== ms) {
      var a = op.anchor, from = Math.min(a[0], ms), to = Math.max(a[1], me);
      op.anchor = null; setDate(k, { from: from, to: to, label: fmtMonthRange(from, to) });
    } else { op.anchor = [ms, me]; setDate(k, { from: ms, to: me, label: 'T' + (m + 1) + '/' + y }); }
    return;
  }
  if (b.dataset.q != null) {
    var yy = op.year;
    if (b.dataset.q === 'y') setDate(k, { from: Date.UTC(yy, 0, 1), to: Date.UTC(yy, 11, 31), label: 'Năm ' + yy });
    else { var q = +b.dataset.q; setDate(k, { from: monthStart(yy, q * 3), to: monthEnd(yy, q * 3 + 2), label: 'Q' + (q + 1) + '/' + yy }); }
    return;
  }
  if (b.dataset.a === 'clear') { op.anchor = null; setDate(k, null); return; }
  if (b.dataset.a === 'close') closePop();
}
function fmtMonthRange(f, t) { var a = new Date(f), b = new Date(t); return 'T' + (a.getUTCMonth() + 1) + '/' + a.getUTCFullYear() + ' → T' + (b.getUTCMonth() + 1) + '/' + b.getUTCFullYear(); }

/* ---- popover cột ---- */
function openColPop(host) {
  if (openPop && openPop.el === host) { closePop(); return; }
  closePop();
  var pop = document.createElement('div'); pop.className = 'pop right colpop';
  host.appendChild(pop); openPop = { el: host, kind: 'cols' };
  function fill() {
    pop.innerHTML = '<h5>Bộ cột có sẵn</h5><div class="presets">' + Object.keys(COL_PRESETS).map(function (p) { return '<button data-cp="' + esc(p) + '">' + esc(p) + '</button>'; }).join('') + '</div>' +
      '<h5 style="margin-top:10px">Bật/tắt từng cột (' + visCols.length + '/' + COLS.length + ')</h5><div class="colgrid">' +
      COLS.map(function (c) { return '<label class="opt"><input type="checkbox" data-ck="' + c.k + '"' + (visCols.indexOf(c.k) >= 0 ? ' checked' : '') + '><span class="tx">' + esc(c.label) + '</span></label>'; }).join('') + '</div>' +
      '<div class="pf"><span style="font-size:12px;color:var(--muted);align-self:center">Lựa chọn được ghi nhớ cho lần sau</span><button class="btn sm primary" data-a="close">Xong</button></div>';
  }
  fill();
  pop.addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    if (b.dataset.a === 'close') { closePop(); return; }
    if (b.dataset.cp) { visCols = COL_PRESETS[b.dataset.cp].slice(); store.set('cols', visCols); renderTable(); fill(); }
  });
  pop.addEventListener('change', function (e) {
    var k = e.target.dataset.ck; if (!k) return;
    if (e.target.checked) { var order = COLS.map(function (c) { return c.k; }); visCols.push(k); visCols.sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); }); }
    else visCols = visCols.filter(function (x) { return x !== k; });
    store.set('cols', visCols); renderTable(); $('h5:nth-of-type(2)', pop).textContent = 'Bật/tắt từng cột (' + visCols.length + '/' + COLS.length + ')';
  });
}

/* ---- bộ lọc đã lưu ---- */
function serializeFilters() {
  var o = { q: F.q, quick: Array.from(F.quick), facets: {}, dates: {} };
  for (var k in F.facets) if (F.facets[k].size) o.facets[k] = Array.from(F.facets[k]);
  for (var d in F.dates) if (F.dates[d]) o.dates[d] = F.dates[d];
  return o;
}
function loadFilters(o) {
  resetFilters(); F.q = o.q || ''; $('#q').value = F.q;
  (o.quick || []).forEach(function (q) { if (QUICK_BY[q]) F.quick.add(q); });
  Object.keys(o.facets || {}).forEach(function (k) { if (F.facets[k]) o.facets[k].forEach(function (v) { F.facets[k].add(v); }); });
  Object.keys(o.dates || {}).forEach(function (k) { if (DATE_BY[k]) F.dates[k] = o.dates[k]; });
  S.page = 1; update();
}
function openViewsPop(host) {
  if (openPop && openPop.el === host) { closePop(); return; }
  closePop();
  var pop = document.createElement('div'); pop.className = 'pop'; host.appendChild(pop); placePop(pop, host); openPop = { el: host, kind: 'views' };
  function fill() {
    var V = store.get('views', []);
    pop.innerHTML = '<h5>Lưu bộ lọc hiện tại</h5><div style="display:flex;gap:6px"><input type="text" id="vName" placeholder="VD: Kaspersky đang xử lý" style="flex:1;height:32px;border:1px solid var(--line-2);border-radius:8px;padding:0 8px;background:var(--card-2)"><button class="btn sm primary" data-a="save">' + icon('bookmark') + 'Lưu</button></div>' +
      '<h5 style="margin-top:12px">Đã lưu (' + V.length + ')</h5>' + (V.length ? V.map(function (v, i) { return '<label class="opt" style="cursor:default"><span class="tx" style="cursor:pointer" data-load="' + i + '">' + esc(v.name) + '</span><button class="btn sm icon ghost" data-del="' + i + '" title="Xóa">' + icon('trash') + '</button></label>'; }).join('') : '<div class="empty">Chưa có bộ lọc nào được lưu</div>');
  }
  fill();
  pop.addEventListener('click', function (e) {
    var t = e.target.closest('[data-a],[data-load],[data-del]'); if (!t) return;
    var V = store.get('views', []);
    if (t.dataset.a === 'save') { var n = $('#vName').value.trim(); if (!n) { toast('Nhập tên bộ lọc', true); return; } if (!activeCount()) { toast('Chưa có bộ lọc nào đang áp dụng', true); return; } V.push({ name: n, f: serializeFilters() }); store.set('views', V); toast('Đã lưu bộ lọc “' + n + '”'); fill(); }
    else if (t.dataset.load != null) { loadFilters(V[+t.dataset.load].f); closePop(); toast('Đã áp dụng “' + V[+t.dataset.load].name + '”'); }
    else if (t.dataset.del != null) { V.splice(+t.dataset.del, 1); store.set('views', V); fill(); }
  });
}

/* ----------------------------- autocomplete tìm kiếm ----------------------------- */
var acState = { items: [], idx: -1 };
function closeAc() { var a = $('#ac'); if (a) a.classList.add('hidden'); acState.idx = -1; }
function buildAc(q) {
  var nq = nk(q), box = $('#ac');
  if (!nq) { closeAc(); return; }
  var toks = nq.split(' ').filter(Boolean);
  function match(s) { var n = nk(s); for (var i = 0; i < toks.length; i++) if (n.indexOf(toks[i]) < 0) return false; return true; }
  var groups = [['kh', 'Khách hàng'], ['hang', 'Hãng'], ['sale', 'Sale'], ['dl', 'Đại lý / Reseller'], ['pm', 'PM thực hiện'], ['nhom', 'Nhóm KH'], ['step', 'Bước hiện tại']];
  var items = [], h = '';
  groups.forEach(function (g) {
    var m = countBy(S.rows, function (r) { return r[g[0]]; }), arr = [];
    m.forEach(function (c, v) { if (v !== EMPTY && match(v)) arr.push([v, c]); });
    arr.sort(function (a, b) { return b[1] - a[1]; });
    if (!arr.length) return;
    h += '<div class="grp">' + g[1] + ' · ' + arr.length + ' kết quả</div>';
    arr.slice(0, g[0] === 'kh' ? 7 : 5).forEach(function (e) { items.push({ t: 'f', k: g[0], v: e[0] }); h += '<div class="it" data-ai="' + (items.length - 1) + '"><span class="tg">' + esc(g[1].split(' ')[0]) + '</span><span class="tx">' + hl(e[0], q) + '</span><span class="ct">' + e[1] + ' phiếu</span></div>'; });
  });
  // Phiếu khớp toàn văn (ID, tên, KH, hãng, sale, đại lý, PM, mô tả, BOM…) — ưu tiên khớp ID/tên, mới nhất trước
  var hits = S.rows.filter(function (r) { for (var i = 0; i < toks.length; i++) if (r.hay.indexOf(toks[i]) < 0) return false; return true; });
  hits.sort(function (a, b) { return (match(b.id + ' ' + b.name) ? 1 : 0) - (match(a.id + ' ' + a.name) ? 1 : 0) || (b.dTao || 0) - (a.dTao || 0); });
  if (hits.length) {
    h += '<div class="grp">Phiếu liên quan · ' + hits.length + ' kết quả</div>';
    hits.slice(0, 12).forEach(function (r) {
      items.push({ t: 'r', r: r });
      h += '<div class="it" data-ai="' + (items.length - 1) + '"><span class="tg">#' + esc(r.id) + '</span><span class="tx">' + hl(r.name, q) +
        '<div style="font-size:11.5px;color:var(--muted);overflow:hidden;text-overflow:ellipsis">' + hl(r.hang + ' · ' + r.kh + ' · ' + r.sale + (r.dl !== EMPTY ? ' · ' + r.dl : ''), q) + '</div></span><span class="ct">' + stBadge(r.st) + '</span></div>';
    });
    items.push({ t: 'all' });
    h += '<div class="it" data-ai="' + (items.length - 1) + '" style="justify-content:center;font-weight:700;color:var(--brand-3)">' + icon('list') + 'Xem tất cả ' + hits.length + ' phiếu trong bảng chi tiết</div>';
  } else h += '<div class="foot">Không có phiếu nào chứa “' + esc(q) + '”.</div>';
  h += '<div class="foot">↑↓ chọn · Enter: mở/áp dụng · Esc: đóng — bảng & biểu đồ đang lọc theo từ khóa “' + esc(q) + '” (' + hits.length + ' phiếu)</div>';
  acState.items = items; acState.idx = -1;
  box.innerHTML = h; box.classList.remove('hidden');
}
function pickAc(i) {
  var it = acState.items[i]; if (!it) return;
  if (it.t === 'f') { F.q = ''; $('#q').value = ''; $('#qClr').classList.add('hidden'); F.facets[it.k].add(it.v); S.page = 1; update(); toast('Đã lọc ' + FACET_BY[it.k].label + ': ' + it.v); }
  else if (it.t === 'all') { closeAc(); $('#tableCard').scrollIntoView({ behavior: 'smooth' }); return; }
  else openDrawer(it.r);
  closeAc();
}

/* ----------------------------- drawer chi tiết ----------------------------- */
function closeLayer() { $('#layer').innerHTML = ''; }
function openDrawer(r) {
  var tl = [['Tạo phiếu', r.dTao, true], ['Cập nhật cuối', r.dCn, true], ['Giao hàng', r.dGiao, r.st === 'Hoàn thành' || (r.dGiao != null && r.dGiao <= S.refDay)], ['License start', r.dLs, r.dLs != null && r.dLs <= S.refDay], ['License end', r.dLe, r.dLe != null && r.dLe <= S.refDay]];
  var kv = [['Hãng', r.hang], ['Khách hàng', r.kh + (r.khSrc !== 'Form khách hàng' ? ' (' + r.khSrc + ')' : '')], ['Nhóm KH', r.nhom], ['Đại lý / Reseller', r.dl], ['Sale', r.sale], ['Người tạo', r.creator], ['PM thực hiện', r.pm],
    ['Luồng nhập hàng', r.flow], ['Tình trạng HĐ', hdBadge(r.hd), 1], ['Công ty ký HĐ', r.cty], ['BB nghiệm thu/thanh lý', r.bb], ['Bước hiện tại', r.step], ['Trạng thái', stBadge(r.st), 1],
    ['Cờ nhắc', bdg(r.flagLabel, FLAG_META[r.flag].b) + (r.idle != null ? ' · ' + r.idle + ' ngày chưa cập nhật' : ''), 1], ['Giao hàng', bdg(r.giaoLabel, GIAO_META[r.giao][2]) + (r.daysToGiao != null && r.st !== 'Hoàn thành' ? ' · ' + (r.daysToGiao < 0 ? 'trễ ' + (-r.daysToGiao) : 'còn ' + r.daysToGiao) + ' ngày' : ''), 1],
    ['License', bdg(r.licLabel, LIC_META[r.lic][2]) + (r.licMonths != null ? ' · thời hạn ' + fmtN(r.licMonths, 1) + ' tháng' : ''), 1], ['TG xử lý', r.tg != null ? fmtN(r.tg, 1) + ' ngày' : '–'], ['Tuổi phiếu', r.age != null ? r.age + ' ngày' : '–'],
    ['SLA/Hạn', r.dSla != null ? fmtDT(r.dSla) + (r.qh === 'Có' ? ' · ' + bdg('Quá hạn', 'b-purple') : '') : '–', 1], ['Cờ review', r.review || '–'], ['Đề xuất hành động', r.action ? '<b style="color:var(--accent)">' + esc(r.action) + '</b>' : '–', 1]];
  $('#layer').innerHTML = '<div class="overlay" data-close></div><div class="drawer"><div class="dh"><button class="btn sm icon x" data-close>' + icon('x') + '</button><div class="id">Phiếu #' + esc(r.id) + ' · No. ' + esc(r.no) + '</div><h3>' + esc(r.name) + '</h3>' + stBadge(r.st) + ' ' + bdg(r.flagLabel, FLAG_META[r.flag].b) + '</div><div class="db">' +
    '<div class="tl">' + tl.map(function (p) { return '<div class="p ' + (p[1] == null ? '' : p[2] ? 'done' : '') + (p[0] === 'Giao hàng' && r.giao === 'late' ? ' late' : '') + '"><i></i><b>' + (p[1] != null ? fmtD(p[1]) : '–') + '</b>' + p[0] + '</div>'; }).join('') + '</div>' +
    '<div class="kv">' + kv.map(function (x) { return '<div>' + esc(x[0]) + '</div><div>' + (x[2] ? x[1] : esc(x[1] === EMPTY ? '–' : x[1])) + '</div>'; }).join('') + '</div>' +
    '<h4 style="margin:16px 0 6px;color:var(--brand)">Mô tả phiếu</h4><div class="txtblock">' + esc(r.mota || '–') + '</div>' +
    '<h4 style="margin:16px 0 6px;color:var(--brand)">Thông tin hàng hóa (BOM)</h4><div class="txtblock">' + esc(r.bom || '–') + '</div></div></div>';
}

/* ----------------------------- modal nhắc việc ----------------------------- */
function modal(title, body, foot, wide) {
  $('#layer').innerHTML = '<div class="overlay" data-close></div><div class="modal"' + (wide ? ' style="width:min(1000px,96vw)"' : '') + '><div class="mh"><h3>' + title + '</h3><button class="btn sm icon ghost x" data-close>' + icon('x') + '</button></div><div class="mb">' + body + '</div><div class="mf">' + foot + '</div></div>';
  return $('#layer .modal');
}
function openReminder() {
  var md = modal(icon('mail') + ' Soạn nội dung nhắc cập nhật',
    '<div class="frm"><span>Nhóm theo</span><div class="seg" id="rmBy"><button data-v="pm" class="on">PM thực hiện</button><button data-v="sale">Sale</button></div>' +
    '<span>Mức cờ tối thiểu</span><div class="seg" id="rmLv"><button data-v="m1" class="on">≥ Mức 1</button><button data-v="m2">≥ Mức 2</button><button data-v="m3">Mức 3</button></div>' +
    '<span>Bao gồm</span><div><label class="chk" style="margin:0"><input type="checkbox" id="rmLate" checked> Phiếu trễ hạn giao hàng</label><label class="chk"><input type="checkbox" id="rmHd"> Phiếu hoàn thành nhưng HĐ chưa ký / chưa trả</label></div></div>' +
    '<p style="font-size:12.5px;color:var(--muted)">Áp dụng trên <b>' + S.filtered.length + '</b> phiếu đang lọc. Nội dung có thể chỉnh sửa trước khi sao chép.</p><textarea id="rmText"></textarea>',
    '<button class="btn" data-close>Đóng</button><button class="btn primary" id="rmCopy">' + icon('copy') + 'Sao chép nội dung</button>');
  var by = 'pm', lv = 'm1';
  function gen() {
    var lvl = { m1: ['m1', 'm2', 'm3'], m2: ['m2', 'm3'], m3: ['m3'] }[lv];
    var late = $('#rmLate').checked, hd = $('#rmHd').checked;
    var R = S.filtered.filter(function (r) { return lvl.indexOf(r.flag) >= 0 || (late && r.giao === 'late') || (hd && r.st === 'Hoàn thành' && !r.hdOk); });
    var g = {}; R.forEach(function (r) { (g[r[by]] = g[r[by]] || []).push(r); });
    var keys = Object.keys(g).sort(function (a, b) { return g[b].length - g[a].length; });
    var t = 'NHẮC CẬP NHẬT TIẾN ĐỘ PHIẾU YÊU CẦU ĐẶT HÀNG (tính đến ' + fmtD(S.ref) + ')\nTổng: ' + R.length + ' phiếu cần xử lý\n';
    keys.forEach(function (k) {
      t += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nKính gửi anh/chị ' + (k === EMPTY ? '(chưa rõ người phụ trách)' : k) + ',\nCác phiếu sau cần được cập nhật/xử lý (' + g[k].length + ' phiếu):\n';
      g[k].sort(function (a, b) { return (b.idle || 0) - (a.idle || 0); }).forEach(function (r, i) {
        var why = [];
        if (r.flag !== 'ok' && r.flag !== 'done') why.push(r.idle + ' ngày chưa cập nhật (' + r.flagLabel.split(' ·')[0] + ')');
        if (r.giao === 'late') why.push('trễ hạn giao ' + (-r.daysToGiao) + ' ngày (' + fmtD(r.dGiao) + ')');
        if (r.st === 'Hoàn thành' && !r.hdOk) why.push('HĐ: ' + r.hd);
        t += (i + 1) + '. #' + r.id + ' – ' + r.name + '\n   Hãng: ' + r.hang + ' | KH: ' + r.kh + ' | ' + (by === 'pm' ? 'Sale: ' + r.sale : 'PM: ' + r.pm) + '\n   Bước: ' + r.step + ' | Cập nhật cuối: ' + fmtDT(r.dCn) + '\n   ⚠ ' + why.join('; ') + '\n';
      });
    });
    t += '\nTrân trọng,\nPhòng Quản lý sản phẩm – NTS Hanoi Corp.';
    $('#rmText').value = t;
  }
  md.addEventListener('click', function (e) {
    var b = e.target.closest('.seg button'); if (!b) return;
    var seg = b.parentNode; $$('button', seg).forEach(function (x) { x.classList.toggle('on', x === b); });
    if (seg.id === 'rmBy') by = b.dataset.v; else lv = b.dataset.v; gen();
  });
  $('#rmLate').onchange = gen; $('#rmHd').onchange = gen;
  $('#rmCopy').onclick = function () { copyText($('#rmText').value); };
  gen();
}
function copyText(t) {
  function fb() { var ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast('Đã sao chép vào clipboard'); } catch (e) { toast('Không sao chép được', true); } ta.remove(); }
  if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(function () { toast('Đã sao chép vào clipboard'); }, fb); else fb();
}

/* ----------------------------- thiết lập ----------------------------- */
function openSettings() {
  var md = modal(icon('sliders') + ' Thiết lập cờ nhắc & phân nhóm khách hàng',
    '<div class="frm">' +
    '<span>Cờ nhắc – Bình thường đến</span><div><input type="number" id="sM1" min="1" value="' + CFG.m1 + '" style="width:90px"> ngày không cập nhật</div>' +
    '<span>Mức 1 (cần theo dõi) đến</span><div><input type="number" id="sM2" min="2" value="' + CFG.m2 + '" style="width:90px"> ngày</div>' +
    '<span>Mức 2 (chậm cập nhật) đến</span><div><input type="number" id="sM3" min="3" value="' + CFG.m3 + '" style="width:90px"> ngày · lớn hơn là <b>Mức 3 – tồn đọng</b></div>' +
    '<span>“Sắp đến hạn giao”</span><div>trong vòng <input type="number" id="sG" min="1" value="' + CFG.giaoSoon + '" style="width:90px"> ngày</div>' +
    '<span>“License sắp hết hạn”</span><div>trong vòng <input type="number" id="sL" min="1" value="' + CFG.licSoon + '" style="width:90px"> ngày (cơ hội renew)</div>' +
    '</div><h4 style="margin:18px 0 6px;color:var(--brand)">Quy tắc phân nhóm khách hàng</h4><p style="font-size:12.5px;color:var(--muted);margin:0 0 6px">Mỗi dòng: <code>Tên nhóm | từ khóa 1, từ khóa 2, …</code>. Xét lần lượt từ trên xuống, khớp không phân biệt dấu/hoa thường; thêm <code>^</code> trước từ khóa để bắt buộc ở đầu tên.</p><textarea id="sRules" style="min-height:220px">' + esc(CFG.rules) + '</textarea>',
    '<button class="btn" id="sReset">Khôi phục mặc định</button><button class="btn" data-close>Hủy</button><button class="btn primary" id="sSave">' + icon('check') + 'Áp dụng</button>');
  $('#sReset').onclick = function () { $('#sM1').value = CFG_DEF.m1; $('#sM2').value = CFG_DEF.m2; $('#sM3').value = CFG_DEF.m3; $('#sG').value = CFG_DEF.giaoSoon; $('#sL').value = CFG_DEF.licSoon; $('#sRules').value = CFG_DEF.rules; };
  $('#sSave').onclick = function () {
    var m1 = +$('#sM1').value, m2 = +$('#sM2').value, m3 = +$('#sM3').value;
    if (!(m1 > 0 && m2 > m1 && m3 > m2)) { toast('Ngưỡng phải tăng dần: Bình thường < Mức 1 < Mức 2', true); return; }
    // giữ bộ lọc cờ theo khóa mức (nhãn thay đổi theo ngưỡng)
    var oldFL = flagLabels(), selF = Array.from(F.facets.flagLabel).map(function (l) { for (var k in oldFL) if (oldFL[k] === l) return k; });
    CFG = { m1: m1, m2: m2, m3: m3, giaoSoon: +$('#sG').value || 7, licSoon: +$('#sL').value || 90, rules: $('#sRules').value };
    store.set('cfg', CFG); closeLayer();
    var nfl = flagLabels(); F.facets.flagLabel = new Set(selF.filter(Boolean).map(function (k) { return nfl[k]; }));
    F.facets.giaoLabel.clear(); F.facets.licLabel.clear(); F.facets.nhom.clear();
    if (S.rows.length) { derive(); renderAll(); }
    toast('Đã áp dụng thiết lập');
  };
}

/* ----------------------------- xuất Excel ----------------------------- */
var XS = {
  title: { font: { bold: true, sz: 16, color: { rgb: '0959A2' } } },
  sub: { font: { italic: true, sz: 10, color: { rgb: '64748B' } } },
  h: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '0959A2' } }, alignment: { wrapText: true, vertical: 'center', horizontal: 'center' }, border: bd('9DB7D5') },
  c: { border: bd('D9E2EC'), alignment: { vertical: 'top' } },
  cw: { border: bd('D9E2EC'), alignment: { vertical: 'top', wrapText: true } },
  tot: { font: { bold: true }, fill: { patternType: 'solid', fgColor: { rgb: 'E8F0FA' } }, border: bd('9DB7D5') },
  kpiK: { font: { bold: true, color: { rgb: '3D5170' } }, fill: { patternType: 'solid', fgColor: { rgb: 'EEF4FC' } }, border: bd('D9E2EC') },
  kpiV: { font: { bold: true, sz: 12, color: { rgb: '0959A2' } }, border: bd('D9E2EC'), alignment: { horizontal: 'left' } }
};
function bd(c) { var s = { style: 'thin', color: { rgb: c } }; return { top: s, bottom: s, left: s, right: s }; }
function fillOf(hex) { return { patternType: 'solid', fgColor: { rgb: hex.replace('#', '') } }; }
var FLAG_FILL = { m1: 'FEF3C7', m2: 'FFEDD5', m3: 'FEE2E2', ok: 'DCFCE7', done: 'F1F5F9' };
// Mỗi ô 1 object style riêng (xlsx-js-style ghi numFmt vào object style → không được dùng chung)
function st(style, z) { var o = JSON.parse(JSON.stringify(style || {})); o.numFmt = z || 'General'; return o; }
function xlCell(v, style, z) {
  if (v == null || v === '' || v === EMPTY) return { t: 's', v: '', s: st(style) };
  if (typeof v === 'number') return { t: 'n', v: v, s: st(style, z), z: z || 'General' };
  return { t: 's', v: String(v), s: st(style) };
}
function dateCell(t, withTime, style) { var z = withTime ? 'dd/mm/yyyy hh:mm' : 'dd/mm/yyyy'; return t == null ? { t: 's', v: '', s: st(style) } : { t: 'n', v: C.toSerial(t), z: z, s: st(style, z) }; }
function sheetFrom(title, sub, headers, rows, widths, opts) {
  opts = opts || {};
  var aoa = [[{ t: 's', v: title, s: st(XS.title) }], [{ t: 's', v: sub, s: st(XS.sub) }], []];
  aoa.push(headers.map(function (h) { return { t: 's', v: h, s: st(XS.h) }; }));
  rows.forEach(function (r) { aoa.push(r); });
  if (opts.total) aoa.push(opts.total);
  var ws = X.utils.aoa_to_sheet(aoa);
  ws['!cols'] = widths.map(function (w) { return { wch: w }; });
  ws['!rows'] = [{ hpt: 24 }, null, null, { hpt: 34 }];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(3, headers.length - 1) } }, { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(3, headers.length - 1) } }];
  if (rows.length) ws['!autofilter'] = { ref: X.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: 3 + rows.length, c: headers.length - 1 } }) };
  return ws;
}
function filterDescText() {
  var t = [];
  if (F.q) t.push('Từ khóa: ' + F.q);
  F.quick.forEach(function (id) { t.push(QUICK_BY[id].label); });
  FACETS.forEach(function (f) { var s = F.facets[f.k]; if (s.size) t.push(f.label + ': ' + Array.from(s).map(function (v) { return f.k === 'month' ? ymLabel(v) : v; }).join(', ')); });
  DATES.forEach(function (d) { if (F.dates[d.k]) t.push(d.label + ': ' + dateDesc(F.dates[d.k])); });
  return t.length ? t.join(' | ') : 'Không áp dụng bộ lọc (toàn bộ dữ liệu)';
}
function defaultExportName() {
  var d = F.dates.dTao, t = S.ref;
  if (d && d.from != null && d.to != null) { var a = new Date(d.from), b = new Date(d.to); if (a.getUTCFullYear() === b.getUTCFullYear() && a.getUTCMonth() === b.getUTCMonth()) t = d.from; }
  else if (F.facets.month.size === 1) { var p = Array.from(F.facets.month)[0].split('-'); t = Date.UTC(+p[0], +p[1] - 1, 1); }
  var dt = new Date(t);
  return 'Tổng hợp PO_Hãng_Tháng ' + (dt.getUTCMonth() + 1) + '.' + dt.getUTCFullYear() + '.xlsx';
}
function openExport() {
  var SH = [['tq', 'Tổng quan (KPI, bộ lọc, nhận định)', 1], ['ct', 'Chi tiết PO (dữ liệu sau lọc + cột phân tích)', 1], ['hang', 'Theo Hãng', 1], ['sale', 'Theo Sale', 1], ['dl', 'Theo Đại lý / Reseller', 1],
    ['kh', 'Theo Nhóm khách hàng', 1], ['thang', 'Hãng × Tháng tạo', 1], ['nhac', 'Cờ nhắc – tồn đọng', 1], ['tt', 'Tuân thủ HĐ & BB thanh lý', 1], ['renew', 'License sắp hết hạn (renew)', 1], ['raw', 'Data_sạch (toàn bộ template chuẩn)', 1]];
  var md = modal(icon('download') + ' Xuất báo cáo Excel',
    '<div class="frm"><span>Tên file</span><input type="text" id="exName" value="' + esc(defaultExportName()) + '">' +
    '<span>Phạm vi dữ liệu</span><div class="seg" id="exScope"><button data-v="f" class="on">Theo bộ lọc (' + S.filtered.length + ' phiếu)</button><button data-v="a">Toàn bộ (' + S.rows.length + ' phiếu)</button></div>' +
    '<span>Cột sheet Chi tiết</span><div class="seg" id="exCols"><button data-v="all" class="on">Tất cả ' + COLS.length + ' cột</button><button data-v="vis">Chỉ cột đang hiển thị (' + visCols.length + ')</button></div>' +
    '<span style="align-self:start">Các sheet</span><div>' + SH.map(function (s) { return '<label class="chk" style="margin:0 0 4px"><input type="checkbox" data-sh="' + s[0] + '"' + (s[2] ? ' checked' : '') + '> ' + s[1] + '</label>'; }).join('') + '</div></div>',
    '<button class="btn" data-close>Hủy</button><button class="btn primary" id="exGo">' + icon('download') + 'Tải file Excel</button>');
  var scope = 'f', colsMode = 'all';
  md.addEventListener('click', function (e) {
    var b = e.target.closest('.seg button'); if (!b) return;
    var seg = b.parentNode; $$('button', seg).forEach(function (x) { x.classList.toggle('on', x === b); });
    if (seg.id === 'exScope') scope = b.dataset.v; else colsMode = b.dataset.v;
  });
  $('#exGo').onclick = function () {
    var name = $('#exName').value.trim() || defaultExportName(); if (!/\.xlsx$/i.test(name)) name += '.xlsx';
    var sheets = $$('input[data-sh]', md).filter(function (i) { return i.checked; }).map(function (i) { return i.dataset.sh; });
    if (!sheets.length) { toast('Chọn ít nhất 1 sheet', true); return; }
    loading(true, 'Đang tạo file Excel…');
    setTimeout(function () {
      try { download(buildReport(scope === 'f' ? S.filtered : S.rows, sheets, colsMode, scope), name); closeLayer(); toast('Đã xuất “' + name + '”'); }
      catch (err) { console.error(err); toast('Lỗi xuất Excel: ' + err.message, true); }
      loading(false);
    }, 30);
  };
}
function buildReport(R, sheets, colsMode, scope) {
  var wb = X.utils.book_new(), freeze = {}, sub = 'NTS Hanoi Corp. · Phòng Quản lý sản phẩm · Ngày đối chiếu ' + fmtDT(S.ref) + ' · Nguồn: ' + S.fileName;
  function add(name, ws, fr) { X.utils.book_append_sheet(wb, ws, name); if (fr) freeze[name] = fr; }
  var done = R.filter(function (r) { return r.st === 'Hoàn thành'; }), ip = R.filter(function (r) { return r.st !== 'Hoàn thành'; });
  var tg = done.map(function (r) { return r.tg; }).filter(function (v) { return v != null; });
  if (sheets.indexOf('tq') >= 0) {
    var hdOk = done.filter(function (r) { return r.hdOk; }).length;
    var kp = [['Phạm vi', scope === 'f' ? 'Theo bộ lọc' : 'Toàn bộ dữ liệu'], ['Bộ lọc áp dụng', scope === 'f' ? filterDescText() : '—'], ['Tổng số phiếu', R.length], ['Đang xử lý', ip.length], ['Hoàn thành', done.length],
      ['Tỷ lệ hoàn thành', R.length ? done.length / R.length : 0, '0.0%'], ['TG xử lý trung bình (ngày)', mean(tg), '0.0'], ['TG xử lý trung vị (ngày)', median(tg), '0.0'],
      ['Cần nhắc cập nhật (> ' + CFG.m1 + ' ngày)', ip.filter(function (r) { return r.flag !== 'ok'; }).length], ['  · Mức 3 – tồn đọng (> ' + CFG.m3 + ' ngày)', ip.filter(function (r) { return r.flag === 'm3'; }).length],
      ['Trễ hạn giao hàng', ip.filter(function (r) { return r.giao === 'late'; }).length], ['Quá hạn SLA', R.filter(function (r) { return r.qh === 'Có'; }).length],
      ['Tỷ lệ tuân thủ ký HĐ (trên phiếu HT)', done.length ? hdOk / done.length : 0, '0.0%'], ['HT nhưng HĐ chưa ký', done.filter(function (r) { return r.hd === 'Chưa ký'; }).length], ['HT nhưng KH chưa trả HĐ', done.filter(function (r) { return r.hd === 'KH đã ký nhưng chưa trả Hợp đồng'; }).length],
      ['BB thanh lý “Có” (trên phiếu HT)', done.filter(function (r) { return r.bb === 'Có'; }).length], ['Exwork', R.filter(function (r) { return r.flow === 'Exwork'; }).length], ['License hết hạn ≤ ' + CFG.licSoon + ' ngày', R.filter(function (r) { return r.lic === 'soon'; }).length],
      ['Số hãng', new Set(R.map(function (r) { return r.hang; })).size], ['Số sale', new Set(R.map(function (r) { return r.sale; })).size], ['Số khách hàng', new Set(R.map(function (r) { return r.kh; })).size], ['Số đại lý', new Set(R.filter(function (r) { return r.dl !== EMPTY; }).map(function (r) { return r.dl; })).size]];
    var rows = kp.map(function (k) { return [{ t: 's', v: k[0], s: st(XS.kpiK) }, xlCell(k[1], Object.assign({}, XS.kpiV, k[0] === 'Bộ lọc áp dụng' ? { alignment: { wrapText: true, vertical: 'top' }, font: { sz: 10 } } : {}), k[2])]; });
    rows.push([], [{ t: 's', v: 'NHẬN ĐỊNH NHANH', s: st(XS.h) }, { t: 's', v: '', s: st(XS.h) }]);
    if (scope === 'f' || !activeCount()) $$('#insights li').forEach(function (li) { rows.push([{ t: 's', v: '•', s: st(XS.c) }, { t: 's', v: li.textContent, s: st(XS.cw) }]); });
    var ws = sheetFrom('TỔNG HỢP PO / YÊU CẦU ĐẶT HÀNG THEO HÃNG', sub, ['Chỉ tiêu', 'Giá trị'], rows, [42, 120]);
    delete ws['!autofilter'];
    add('Tổng quan', ws);
  }
  if (sheets.indexOf('ct') >= 0) {
    var cols = colsMode === 'vis' ? visCols.map(function (k) { return COL_BY[k]; }) : COLS;
    var drows = R.map(function (r) {
      var ff = FLAG_FILL[r.flag];
      return cols.map(function (c) {
        var st = c.w ? XS.cw : XS.c;
        if (c.k === 'flag') st = Object.assign({}, XS.c, { fill: fillOf(ff) });
        if (c.k === 'giaoLabel' && r.giao === 'late') st = Object.assign({}, XS.c, { fill: fillOf('FEE2E2'), font: { color: { rgb: 'B91C1C' }, bold: true } });
        if (c.k === 'hd' && r.hd === 'Chưa ký') st = Object.assign({}, XS.c, { font: { color: { rgb: 'B91C1C' }, bold: true } });
        if (c.dt) return dateCell(r[c.k], true, st); if (c.d) return dateCell(r[c.k], false, st);
        if (c.k === 'id') return xlCell(r.id, st, '0');
        if (c.k === 'month') return xlCell(r.month === EMPTY ? '' : r.month, st);
        if (c.n) return xlCell(r[c.k], st, c.dec ? '0.0' : '0');
        var v = c.x ? c.x(r) : r[c.k]; return xlCell(v, st);
      });
    });
    var wmap = { name: 46, kh: 38, dl: 30, mota: 50, bom: 40, review: 36, action: 46, step: 24, hd: 26, flag: 20, nhom: 24 };
    add('Chi tiết PO', sheetFrom('CHI TIẾT PHIẾU YÊU CẦU ĐẶT HÀNG (' + R.length + ' phiếu)', sub + ' · ' + (scope === 'f' ? filterDescText() : 'Toàn bộ dữ liệu'), cols.map(function (c) { return c.label; }), drows, cols.map(function (c) { return wmap[c.k] || (c.dt ? 17 : c.d ? 13 : c.n ? 11 : 18); })), 'C5');
  }
  function groupSheet(key, label, sheetName, title) {
    var by = {};
    R.forEach(function (r) {
      var k = r[key], o = by[k] = by[k] || { n: 0, ip: 0, done: 0, tg: [], flag: 0, m3: 0, late: 0, sla: 0, hdOk: 0, hdNo: 0, hdRet: 0, bb: 0, exw: 0, lic: 0, hang: new Set(), kh: new Set() };
      o.n++; if (r.st === 'Hoàn thành') { o.done++; if (r.tg != null) o.tg.push(r.tg); if (r.hdOk) o.hdOk++; if (r.bb === 'Có') o.bb++; } else o.ip++;
      if (r.flag === 'm1' || r.flag === 'm2' || r.flag === 'm3') o.flag++; if (r.flag === 'm3') o.m3++; if (r.giao === 'late') o.late++; if (r.qh === 'Có') o.sla++;
      if (r.hd === 'Chưa ký') o.hdNo++; if (r.hd === 'KH đã ký nhưng chưa trả Hợp đồng') o.hdRet++; if (r.flow === 'Exwork') o.exw++; if (r.lic === 'soon') o.lic++;
      o.hang.add(r.hang); o.kh.add(r.kh);
    });
    var ks = Object.keys(by).sort(function (a, b) { return by[b].n - by[a].n; });
    var H = [label, 'Tổng phiếu', 'Tỷ trọng', 'Đang xử lý', 'Hoàn thành', 'Tỷ lệ HT', 'TG xử lý TB (ngày)', 'TG xử lý trung vị (ngày)', 'Cần nhắc', 'Tồn đọng mức 3', 'Trễ hạn giao', 'Quá hạn SLA', 'HĐ chưa ký', 'KH chưa trả HĐ', 'Tuân thủ HĐ (trên HT)', 'BB thanh lý Có', 'Exwork', 'License sắp hết hạn', key === 'hang' ? 'Số KH' : 'Số hãng', key === 'hang' ? 'Hãng/KH' : 'Số KH'];
    var tot = { n: 0, ip: 0, done: 0, flag: 0, m3: 0, late: 0, sla: 0, hdOk: 0, hdNo: 0, hdRet: 0, bb: 0, exw: 0, lic: 0 };
    var rows = ks.map(function (k) {
      var o = by[k]; for (var f in tot) tot[f] += o[f];
      return [xlCell(k, XS.c), xlCell(o.n, XS.c), xlCell(o.n / R.length, XS.c, '0.0%'), xlCell(o.ip, XS.c), xlCell(o.done, XS.c), xlCell(o.n ? o.done / o.n : 0, XS.c, '0.0%'), xlCell(mean(o.tg), XS.c, '0.0'), xlCell(median(o.tg), XS.c, '0.0'),
        xlCell(o.flag, XS.c), xlCell(o.m3, XS.c), xlCell(o.late, XS.c), xlCell(o.sla, XS.c), xlCell(o.hdNo, XS.c), xlCell(o.hdRet, XS.c), xlCell(o.done ? o.hdOk / o.done : null, XS.c, '0.0%'), xlCell(o.bb, XS.c), xlCell(o.exw, XS.c), xlCell(o.lic, XS.c),
        xlCell(key === 'hang' ? o.kh.size : o.hang.size, XS.c), xlCell(key === 'hang' ? Math.round(o.n / Math.max(1, o.kh.size) * 10) / 10 : o.kh.size, XS.c)];
    });
    var T = [xlCell('TỔNG', XS.tot), xlCell(tot.n, XS.tot), xlCell(1, XS.tot, '0.0%'), xlCell(tot.ip, XS.tot), xlCell(tot.done, XS.tot), xlCell(tot.n ? tot.done / tot.n : 0, XS.tot, '0.0%'), xlCell(mean(tg), XS.tot, '0.0'), xlCell(median(tg), XS.tot, '0.0'),
      xlCell(tot.flag, XS.tot), xlCell(tot.m3, XS.tot), xlCell(tot.late, XS.tot), xlCell(tot.sla, XS.tot), xlCell(tot.hdNo, XS.tot), xlCell(tot.hdRet, XS.tot), xlCell(tot.done ? tot.hdOk / tot.done : null, XS.tot, '0.0%'), xlCell(tot.bb, XS.tot), xlCell(tot.exw, XS.tot), xlCell(tot.lic, XS.tot), xlCell('', XS.tot), xlCell('', XS.tot)];
    add(sheetName, sheetFrom(title, sub, H, rows, [key === 'dl' ? 44 : 30, 10, 9, 10, 10, 9, 12, 12, 10, 11, 10, 10, 10, 11, 12, 10, 9, 11, 9, 9], { total: T }), 'B5');
  }
  if (sheets.indexOf('hang') >= 0) groupSheet('hang', 'Hãng', 'Theo Hãng', 'TỔNG HỢP THEO HÃNG');
  if (sheets.indexOf('sale') >= 0) groupSheet('sale', 'Sale', 'Theo Sale', 'TỔNG HỢP THEO SALE (CHỦ SỞ HỮU THƯƠNG VỤ)');
  if (sheets.indexOf('dl') >= 0) groupSheet('dl', 'Đại lý / Reseller', 'Theo Đại lý', 'TỔNG HỢP THEO ĐẠI LÝ / RESELLER');
  if (sheets.indexOf('kh') >= 0) groupSheet('nhom', 'Nhóm khách hàng', 'Theo Nhóm KH', 'TỔNG HỢP THEO NHÓM KHÁCH HÀNG');
  if (sheets.indexOf('thang') >= 0) {
    var months = Array.from(new Set(R.map(function (r) { return r.month; }))).sort(), hs = hangStats(R), cell = {};
    R.forEach(function (r) { var k = r.hang + '|' + r.month; cell[k] = (cell[k] || 0) + 1; });
    var mx = Math.max.apply(null, Object.values(cell).concat([1]));
    var rows2 = hs.map(function (o) {
      return [xlCell(o.k, XS.c)].concat(months.map(function (m) { var v = cell[o.k + '|' + m] || 0; var a = v / mx; var col = a > .66 ? '7FB2E5' : a > .33 ? 'B9D5F2' : v ? 'E3EEFA' : 'FFFFFF'; return xlCell(v || null, Object.assign({}, XS.c, { fill: fillOf(col), alignment: { horizontal: 'center' } })); })).concat([xlCell(o.n, XS.tot)]);
    });
    var T2 = [xlCell('TỔNG', XS.tot)].concat(months.map(function (m) { return xlCell(R.filter(function (r) { return r.month === m; }).length, XS.tot); })).concat([xlCell(R.length, XS.tot)]);
    add('Hãng x Tháng', sheetFrom('MA TRẬN SỐ PHIẾU HÃNG × THÁNG TẠO', sub, ['Hãng'].concat(months.map(ymLabel)).concat(['Tổng']), rows2, [22].concat(months.map(function () { return 10; })).concat([10]), { total: T2 }), 'B5');
  }
  function listSheet(name, title, rows, cols, fr) {
    var data = rows.map(function (r) { var ff = FLAG_FILL[r.flag]; return cols.map(function (k) { var c = COL_BY[k], st = c.w ? XS.cw : XS.c; if (k === 'flag') st = Object.assign({}, XS.c, { fill: fillOf(ff) }); if (c.dt) return dateCell(r[k], true, st); if (c.d) return dateCell(r[k], false, st); if (c.n) return xlCell(r[k], st, c.dec ? '0.0' : '0'); return xlCell(c.x ? c.x(r) : r[k], st); }); });
    var wmap = { name: 46, kh: 36, action: 50, step: 24, hd: 26, flag: 20, dl: 28 };
    add(name, sheetFrom(title + ' (' + rows.length + ' phiếu)', sub, cols.map(function (k) { return COL_BY[k].label; }), data, cols.map(function (k) { var c = COL_BY[k]; return wmap[k] || (c.dt ? 17 : c.d ? 13 : c.n ? 11 : 18); })), fr);
  }
  if (sheets.indexOf('nhac') >= 0) {
    var nh = R.filter(function (r) { return r.st !== 'Hoàn thành' && (r.flag !== 'ok' || r.giao === 'late' || r.qh === 'Có'); }).sort(function (a, b) { return (b.idle || 0) - (a.idle || 0); });
    listSheet('Cờ nhắc', 'DANH SÁCH PHIẾU CẦN NHẮC CẬP NHẬT / TỒN ĐỌNG / TRỄ HẠN', nh, ['id', 'name', 'hang', 'kh', 'sale', 'pm', 'step', 'dTao', 'dCn', 'idle', 'flag', 'dGiao', 'daysToGiao', 'giaoLabel', 'qh', 'action'], 'C5');
  }
  if (sheets.indexOf('tt') >= 0) {
    var tt = R.filter(function (r) { return r.st === 'Hoàn thành' && (!r.hdOk || r.bb !== 'Có'); }).sort(function (a, b) { return (a.hdOk ? 1 : 0) - (b.hdOk ? 1 : 0) || String(a.sale).localeCompare(String(b.sale), 'vi'); });
    listSheet('Tuân thủ HĐ-BB', 'PHIẾU HOÀN THÀNH CHƯA ĐỦ HỢP ĐỒNG / BIÊN BẢN THANH LÝ', tt, ['id', 'name', 'hang', 'kh', 'dl', 'sale', 'hd', 'cty', 'bb', 'hdOkLabel', 'dTao', 'dCn', 'action'], 'C5');
  }
  if (sheets.indexOf('renew') >= 0) {
    var rn = R.filter(function (r) { return r.lic === 'soon' || (r.lic === 'expired' && r.daysToLe >= -90); }).sort(function (a, b) { return a.dLe - b.dLe; });
    listSheet('License renew', 'LICENSE HẾT HẠN TRONG ' + CFG.licSoon + ' NGÀY TỚI / VỪA HẾT HẠN ≤ 90 NGÀY – CƠ HỘI RENEW', rn, ['id', 'name', 'hang', 'kh', 'dl', 'sale', 'dLs', 'dLe', 'daysToLe', 'licMonths', 'licLabel'], 'C5');
  }
  if (sheets.indexOf('raw') >= 0) {
    add(C.SHEET_NAME, C.cleanSheet(X, S.template), 'A2'); // giữ nguyên định dạng template chuẩn
  }
  wb.Props = { Title: 'Tổng hợp PO theo Hãng', Company: 'NTS Hanoi Corp.', Author: 'Phòng Quản lý sản phẩm' };
  return C.freezeAndWrite(X, wb, freeze);
}
function downloadClean() {
  var base = S.fileName.replace(/\.(xlsx|xlsm|xls|csv)$/i, '');
  var name = S.mode === 'clean' ? base + '.xlsx' : base + '_CLEAN.xlsx';
  download(C.writeClean(X, S.template), name);
  toast('Đã tải file template chuẩn “' + name + '” (' + S.template.length + ' dòng, sheet ' + C.SHEET_NAME + ')');
}
function copyTable() {
  var cols = visCols.map(function (k) { return COL_BY[k]; });
  var lines = [cols.map(function (c) { return c.label; }).join('\t')];
  S.filtered.forEach(function (r) { lines.push(cols.map(function (c) { var v = cellText(r, c); return v == null ? '' : String(v).replace(/[\t\n\r]+/g, ' '); }).join('\t')); });
  copyText(lines.join('\n'));
}

/* ----------------------------- giao diện sáng/tối ----------------------------- */
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  $('#btnTheme').innerHTML = icon(t === 'dark' ? 'sun' : 'moon');
  store.set('theme', t);
  chartBase();
  Object.keys(CH).forEach(function (k) { CH[k].destroy(); delete CH[k]; });
  if (S.rows.length) { renderOverview(); renderCharts(); }
}

/* ----------------------------- sự kiện ----------------------------- */
function bindStatic() {
  $('#quick').addEventListener('click', function (e) { var c = e.target.closest('.qc'); if (!c) return; var id = c.dataset.q; if (F.quick.has(id)) F.quick.delete(id); else F.quick.add(id); S.page = 1; update(); });
  $('#fgroups').addEventListener('click', function (e) {
    var b = e.target.closest('.fp > button'); if (!b) return;
    var host = b.parentNode; e.stopPropagation();
    if (host.dataset.fk) openFacetPop(host, host.dataset.fk); else if (host.dataset.dk) openDatePop(host, host.dataset.dk);
  });
  $('#activeBar').addEventListener('click', function (e) { var b = e.target.closest('[data-rm]'); if (b) removeTag(b.dataset.rm); });
  $('#btnClearAll').onclick = function () { resetFilters(); S.page = 1; update(); toast('Đã xóa toàn bộ bộ lọc'); };
  $('#btnViews').onclick = function (e) { e.stopPropagation(); openViewsPop($('#viewsFp')); };
  $('#btnCols').onclick = function (e) { e.stopPropagation(); openColPop($('#colFp')); };
  // tìm kiếm
  var q = $('#q');
  q.addEventListener('input', function () { F.q = q.value.trim(); $('#qClr').classList.toggle('hidden', !q.value); $('#qKbd').classList.toggle('hidden', !!q.value); buildAc(q.value); S.page = 1; updateDebounced(); });
  q.addEventListener('focus', function () { if (q.value) buildAc(q.value); });
  q.addEventListener('keydown', function (e) {
    var its = $$('#ac .it');
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); if (!its.length) return; acState.idx = (acState.idx + (e.key === 'ArrowDown' ? 1 : -1) + its.length) % its.length; its.forEach(function (x, i) { x.classList.toggle('on', i === acState.idx); }); its[acState.idx].scrollIntoView({ block: 'nearest' }); }
    else if (e.key === 'Enter') { if (acState.idx >= 0) pickAc(+its[acState.idx].dataset.ai); else { closeAc(); if (F.q) $('#tableCard').scrollIntoView({ behavior: 'smooth' }); } }
    else if (e.key === 'Escape') closeAc();
  });
  $('#ac').addEventListener('mousedown', function (e) { var it = e.target.closest('.it'); if (it) { e.preventDefault(); pickAc(+it.dataset.ai); } });
  $('#qClr').onclick = function () { q.value = ''; F.q = ''; $('#qClr').classList.add('hidden'); $('#qKbd').classList.remove('hidden'); closeAc(); update(); };
  // ngày đối chiếu
  $('#refMode').onchange = function () { S.refMode = this.value; if (S.refMode === 'custom') { if (S.refCustom == null) S.refCustom = S.refDay; $('#refDate').value = isoD(S.refCustom); } derive(); renderAll(); };
  $('#refDate').onchange = function () { var v = parseIso(this.value); if (v != null) { S.refCustom = v; derive(); renderAll(); } };
  // bảng
  $('#dt').addEventListener('click', function (e) {
    var th = e.target.closest('th[data-sort]');
    if (th) { var k = th.dataset.sort; S.sort = { k: k, d: S.sort.k === k ? -S.sort.d : (COL_BY[k].n || COL_BY[k].d || COL_BY[k].dt ? -1 : 1) }; store.set('sort', S.sort); applyFilters(); renderTable(); return; }
    var a = e.target.closest('[data-rm]'); if (a) { e.preventDefault(); removeTag('*'); return; }
    var tr = e.target.closest('tr[data-i]'); if (tr) openDrawer(S.rows[+tr.dataset.i]);
  });
  $('#pager').addEventListener('click', function (e) { var b = e.target.closest('[data-pg]'); if (!b || b.disabled) return; S.page = +b.dataset.pg; renderTable(); $('#tableCard').scrollIntoView({ behavior: 'smooth', block: 'start' }); });
  $('#pageSize').onchange = function () { S.pageSize = +this.value; S.page = 1; renderTable(); };
  $('#btnCopy').onclick = copyTable;
  $('#btnExport2').onclick = openExport;
  $('#dlTop').onchange = renderCharts;
  // seg toggles
  $$('.seg[data-seg]').forEach(function (seg) {
    var k = seg.dataset.seg;
    if (k === 'density') $$('button', seg).forEach(function (b) { b.classList.toggle('on', b.dataset.v === S.density); });
    seg.addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      $$('button', seg).forEach(function (x) { x.classList.toggle('on', x === b); });
      S[k] = b.dataset.v;
      if (k === 'density') { store.set('density', S.density); renderTable(); } else if (k === 'calMode') renderCal(); else renderCharts();
    });
  });
  // click legend/tables để lọc
  document.addEventListener('click', function (e) {
    var t = e.target.closest('[data-f][data-v]');
    if (t && !t.closest('.pop')) { toggleFacet(t.dataset.f, t.dataset.v, true); return; }
    var hm = e.target.closest('[data-hm]');
    if (hm && hm.textContent) { var p = hm.dataset.hm.split('|'); F.facets.hang = new Set([p[0]]); F.facets.month = new Set([p[1]]); S.page = 1; update(); return; }
    var kp = e.target.closest('.kpi[data-act]');
    if (kp) { var id = kp.dataset.act; if (F.quick.has(id)) F.quick.delete(id); else F.quick.add(id); S.page = 1; update(); }
  });
}
function bindGlobal() {
  var drop = $('#drop'), fi = $('#fileInput');
  drop.addEventListener('click', function () { fi.click(); });
  fi.addEventListener('change', function () { handleFile(fi.files[0]); fi.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { document.addEventListener(ev, function (e) { e.preventDefault(); if (ev === 'dragleave' && e.relatedTarget) return; drop.classList.remove('over'); }); });
  document.addEventListener('drop', function (e) { var f = e.dataTransfer && e.dataTransfer.files[0]; if (f) handleFile(f); });
  $('#btnReload').onclick = function () { fi.click(); };
  $('#btnClean').onclick = downloadClean;
  $('#btnExport').onclick = openExport;
  $('#btnSettings').onclick = openSettings;
  $('#btnTheme').onclick = function () { applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'); };
  $('#layer').addEventListener('click', function (e) { if (e.target.closest('[data-close]')) closeLayer(); });
  document.addEventListener('click', function (e) {
    if (openPop && !openPop.el.contains(e.target) && document.body.contains(e.target)) closePop();
    if (!e.target.closest('#searchBox')) closeAc();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { if ($('#layer').innerHTML) closeLayer(); else closePop(); closeAc(); }
    if (e.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) && S.rows.length) { e.preventDefault(); $('#q').focus(); }
  });
}

/* ----------------------------- khởi động ----------------------------- */
function boot() {
  if (!X || !window.Chart || !C) { document.body.insertAdjacentHTML('afterbegin', '<div class="warnbox">Không tải được thư viện nhúng. Vui lòng dùng file HTML đầy đủ.</div>'); return; }
  chartBase();
  applyTheme(store.get('theme', (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light'));
  bindGlobal();
}
// API nội bộ phục vụ kiểm thử tự động
window.NTSApp = { S: S, F: F, loadWorkbook: loadWorkbook, buildReport: buildReport, update: update, toggleFacet: toggleFacet, setDate: setDate };
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
