/* ============================================================================
 * converter.js — Port JavaScript 1:1 của convert_service_tickets.py
 * Chuyển file RAW "Service Tickets - Yêu cầu đặt hàng" (xuất từ Base) thành
 * bảng chuẩn "Data_sạch" (30 cột, đúng cấu trúc bản đã duyệt).
 * Chạy được cả trong trình duyệt (window.NTSConvert) lẫn Node (module.exports).
 * Ghi chú tương thích regex: Python `\w`, `\d`, `\b` mặc định là Unicode còn JS
 * là ASCII → dùng \p{L}\p{N} với cờ "u" để cho kết quả giống hệt Python.
 * ==========================================================================*/
(function (root) {
  'use strict';

  var SHEET_NAME = 'Data_sạch';
  var REQUIRED = ['ID', 'Tên', 'Tên Hãng', 'Tên khách hàng', 'Đại lý', 'Sale', 'Người tạo', 'Người thực hiện', 'Tạo lúc',
    'Cập nhật lần cuối lúc', 'Thời gian giao hàng', 'License start date', 'License end date',
    'Luồng nhập hàng', 'Tình trạng Hợp đồng với khách hàng', 'Trạng thái', 'Tên khối'];
  var OPTIONAL = ['No.', 'Công ty ký HĐ', 'Biên bản nghiệm thu thanh lý', 'SLA/Thời hạn', 'Quá hạn', 'Mô tả phiếu', 'Thông tin hàng hóa (BOM)'];

  // ---- Cấu hình mapping (giống script Python) ----
  var ROSTER0 = ['Vũ Việt Anh', 'Đỗ Tuấn Anh', 'Lê Tuấn Anh', 'Hoàng Việt Anh', 'Phạm Thành Chung', 'Lê Trọng Đại', 'Nguyễn Công Đoàn', 'Phạm Mạnh Hà', 'Nguyễn Thu Hằng', 'Đoàn Thị Hiền', 'Trần Hoàng Hiệp', 'Đinh Văn Hiệu', 'Nguyễn Thị Huệ', 'Lê Huy Hùng', 'Phạm Quốc Hùng', 'Trần Quang Hưng', 'Hà Thị Mai Hương', 'Ngọ Lan Hương', 'Nguyễn Thị Hường', 'Nguyễn Văn Hiệp', 'Mai Văn Khá', 'Tạ Diên Khải', 'Nguyễn Văn Khánh', 'Khổng Đức Kiên', 'Nguyễn Trung Kiên', 'Phí Ngọc Linh', 'Phan Việt Linh', 'Nguyễn Văn Minh', 'Nguyễn Đình Minh', 'Phạm Văn Nam', 'Nguyễn Thị Nga', 'Phan Quế Nghiêm', 'Bùi Minh Ngọc', 'Nguyễn Thị Hồng Nhung', 'Lê Thị Phương', 'Vũ Thị Lan Phương', 'Lý Anh Quốc', 'Vy Công Quý', 'Vũ Thị Thanh Quý', 'Phạm Văn Quyết', 'Vũ Xuân Sơn', 'Nguyễn Mỹ Tâm', 'Phạm Thanh Tâm', 'Phạm Duy Thắng', 'Phùng Văn Thành', 'Hoàng Văn Thảo', 'Dương Văn Thế', 'Đinh Văn Thuỷ', 'Nguyễn Duy Tiến', 'Trần Thị Minh Trâm', 'Đặng Đình Trường', 'Phạm Cao Anh Tú', 'Trần Văn Tuấn', 'Nguyễn Thị Tuyển', 'Nguyễn Đức Phúc Tường', 'Lương Thị Tuyết Trinh', 'Lê Phạm Hà Tân', 'Vũ Ngọc Phương', 'Đỗ Trung Hiếu', 'Trần Thị Mai Anh', 'Nguyễn Thị Thu Thuỷ', 'Khương Anh Khôi', 'Hoàng Thị Hương', 'Nguyễn Thị Thảo Lam', 'Nguyễn Duy Thắng', 'Đồng Văn Khoa', 'Nguyễn Thị Mỹ'];
  var UMAP0 = { huongngo: 'Ngọ Lan Hương', huenguyen: 'Nguyễn Thị Huệ', tramtran: 'Trần Thị Minh Trâm', tuantran: 'Trần Văn Tuấn', tupham: 'Phạm Cao Anh Tú', nhungnguyen: 'Nguyễn Thị Hồng Nhung', thuydinh: 'Đinh Văn Thuỷ', minhnguyen: 'Nguyễn Văn Minh', minhnguyen2: 'Nguyễn Đình Minh', kienkhong: 'Khổng Đức Kiên', nganguyen: 'Nguyễn Thị Nga', anhtranmai: 'Trần Thị Mai Anh', phuongle: 'Lê Thị Phương', huongha: 'Hà Thị Mai Hương', linhphan: 'Phan Việt Linh', hungle: 'Lê Huy Hùng' };
  var VMAP0 = { kaspersky: 'Kaspersky', hilstone: 'Hillstone', infoexpress: 'InfoExpress', inforexpress: 'InfoExpress', barracuda: 'Barracuda', safetica: 'Safetica', opentext: 'OpenText', optswat: 'OPSWAT', gtb: 'GTB', hcl: 'HCL', netscout: 'NetScout', netgear: 'Netgear', cloudflare: 'Cloudflare', progress: 'Progress', zecurion: 'Zecurion', qualys: 'Qualys', acronis: 'Acronis', sophos: 'Sophos', delinea: 'Delinea', 'module fs': 'Module FS', 'owl cyber defense': 'Owl Cyber Defense' };

  // Thứ tự & tên cột đầu ra (giữ nguyên như script Python, kể cả 2 cột "(username)")
  var OUT_COLS = ['No. (gốc)', 'ID phiếu', 'Tên phiếu', 'Hãng', 'Khách hàng (chuẩn hóa)', 'Nguồn tên KH', 'Đại lý', 'Sale', 'Người tạo', 'Người thực hiện',
    'Luồng nhập hàng', 'Tình trạng HĐ với KH', 'Công ty ký HĐ (*)', 'Biên bản nghiệm thu/thanh lý (*)',
    'Ngày tạo', 'Cập nhật lần cuối', 'Tháng tạo', 'Ngày giao hàng', 'License start (*)', 'License end',
    'Bước hiện tại', 'Trạng thái', 'Thời gian xử lý (ngày, phiếu hoàn thành)', 'SLA/Hạn (*)', 'Quá hạn (*)', 'Cờ review', 'Mô tả phiếu', 'BOM',
    'Người tạo (username)', 'Người thực hiện (username)'];
  var COL_WIDTHS = [8, 9, 48, 13, 42, 15, 28, 22, 22, 22, 12, 26, 12, 14, 17, 17, 10, 13, 13, 13, 26, 12, 14, 17, 10, 40, 50, 40, 16, 16];
  var DATETIME_COLS = ['Ngày tạo', 'Cập nhật lần cuối', 'SLA/Hạn (*)'];
  var DATE_COLS = ['Ngày giao hàng', 'License start (*)', 'License end'];

  // ---------------- helpers mô phỏng Python ----------------
  var W = '\\p{L}\\p{N}_'; // = Python \w (Unicode)
  var SP = '\\t\\n\\v\\f\\r \\x1c-\\x1f\\x85\\xa0\\u1680\\u2000-\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000'; // = Python \s (Unicode); JS \s khác ở \ufeff, \x1c-\x1f, \x85
  function rx(src, fl) { return new RegExp(src.replace(/\\s/g, '[' + SP + ']'), fl); }

  function isNA(v) { return v === null || v === undefined || (typeof v === 'number' && isNaN(v)); }
  function isStr(v) { return typeof v === 'string'; }
  var STRIP_RE = rx('^\\s+|\\s+$', 'gu');
  function pyStrip(s) { return s.replace(STRIP_RE, ''); }
  function stripChars(s, chars) {
    var a = 0, b = s.length;
    while (a < b && chars.indexOf(s[a]) >= 0) a++;
    while (b > a && chars.indexOf(s[b - 1]) >= 0) b--;
    return s.slice(a, b);
  }
  var WS_RE = rx('\\s+', 'gu');
  function collapse(s) { return s.replace(WS_RE, ' '); }
  function pyLen(s) { return Array.from(s).length; }
  function pyIsUpper(s) {
    var cased = false;
    for (var ch of s) {
      var lo = ch.toLowerCase(), up = ch.toUpperCase();
      if (lo !== up) { cased = true; if (ch !== up) return false; }
    }
    return cased;
  }
  function nfc(s) { return s.normalize('NFC'); }
  // numpy round (half-to-even) với 1 chữ số thập phân
  function npRound1(x) {
    var y = x * 10, f = Math.floor(y), d = y - f, r;
    if (d > 0.5) r = f + 1; else if (d < 0.5) r = f; else r = (f % 2 === 0) ? f : f + 1;
    return r / 10;
  }

  // ---------------- customer ----------------
  var KEYW = '(tên|company|e\\.?u\\.?|end[- ]?user|sử dụng|bên|đơn vị|name|khách hàng|chủ đầu tư|thông tin)';
  var KEYW_RE = new RegExp(KEYW, 'iu');
  var DROP = rx('^[^' + W + ']*(địa chỉ|address|điện thoại|phone|tel(?![' + W + '])|sđt|người liên hệ|liên hệ|contact|e-?mail|mst|mã số thuế|website)', 'iu');
  var PLACEHOLDER = /(theo file|thể hiện trên|trên license|theo tài liệu|theo hợp đồng|đính kèm|sale bổ sung|thông tin khách hàng trên|chưa có|^-+$|^n\/?a$|^tbd)/iu;
  var LABEL_RE = rx('^([^:]{0,60}):\\s*([^\\n]+)$', 'u');
  var SPLIT_RE = rx('\\s*(?:(?<![' + W + '])(?:địa chỉ|address|điện thoại|phone|người liên hệ|e-?mail)\\s*:)', 'iu');

  var LEAD_RE = new RegExp('^[' + SP + '\\-•·▪\\*\\p{Nd}\\.\\)\\t]+', 'u');
  var PFX_RE = rx('^(ên|ủ đầu tư là|là)\\s*:?\\s*(?=[A-ZÀ-Ỹ])', 'u');
  var EU_RE = rx('EU\\s*:\\s*([^\\n]+)\\n?$', 'iu');
  function cleanLabel(l) {
    l = pyStrip(l.replace(/ /g, ' '));
    l = l.replace(LEAD_RE, '');
    for (var i = 0; i < 2; i++) {
      var m = LABEL_RE.exec(l);
      if (m && KEYW_RE.test(m[1])) l = m[2];
    }
    l = l.replace(PFX_RE, '');
    l = l.split(SPLIT_RE)[0];
    return stripChars(pyStrip(collapse(l)), ' -•:;,.');
  }
  function custText(t) {
    if (!isStr(t)) return null;
    var lines = t.split('\n');
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (!pyStrip(l)) continue;
      var c = cleanLabel(l);
      if (!c || DROP.test(c) || DROP.test(pyStrip(l))) continue;
      return c;
    }
    return null;
  }
  function custName(n) {
    if (!isStr(n)) return null;
    var m = EU_RE.exec(n);
    if (m) return pyStrip(m[1]);
    var p = n.split('_');
    return p.length > 1 ? pyStrip(p[p.length - 1]) : null;
  }
  function pick(a, b) {
    a = isStr(a) ? a : null; b = isStr(b) ? b : null;
    if (a && !PLACEHOLDER.test(a) && pyLen(a) >= 3) return [a, 'Form khách hàng'];
    if (b) return [b, 'Tên phiếu'];
    return [null, 'Không xác định'];
  }

  // ---------------- dates ----------------
  // Lưu ngày giờ dạng "naive" (ms theo UTC của giờ địa phương ghi trong file) để không lệch múi giờ.
  function mkDate(y, mo, d, h, mi) {
    if (mo < 1 || mo > 12 || d < 1 || h > 23 || mi > 59) return null;
    var t = Date.UTC(y, mo - 1, d, h, mi);
    var dt = new Date(t);
    if (dt.getUTCDate() !== d || dt.getUTCMonth() !== mo - 1) return null;
    return t;
  }
  function parseDT(v, withTime) {
    if (isNA(v)) return null;
    if (v && typeof v === 'object' && '__serial' in v) { // ô Excel kiểu ngày thật (bổ sung so với bản Python)
      return Math.round((v.__serial - 25569) * 1440) * 60000;
    }
    if (v instanceof Date) return isNaN(v) ? null : Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes());
    var s = String(v), m;
    if (withTime) {
      m = /^(\d{1,2}):(\d{1,2}) (\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
      return m ? mkDate(+m[5], +m[4], +m[3], +m[1], +m[2]) : null;
    }
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
    return m ? mkDate(+m[3], +m[2], +m[1], 0, 0) : null;
  }
  function ymOf(t) { var d = new Date(t); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); }

  // ---------------- PII mask ----------------
  var EMAIL_RE = new RegExp('[' + W + '.\\-+]+@[' + W + '\\-]+(\\.[' + W + '\\-]+)+', 'gu');
  var PHONE_RE = /(?<!\p{Nd})(?:\+?84|0)[\p{Nd} .\-]{8,12}\p{Nd}/gu;
  function mask(s) {
    if (!isStr(s)) return s;
    s = s.replace(EMAIL_RE, '[email]');
    s = s.replace(PHONE_RE, '[sdt]');
    return pyStrip(collapse(s.replace(/ /g, ' ')));
  }

  function toNumber(v) {
    if (isNA(v)) return null;
    if (typeof v === 'number') return v;
    var s = pyStrip(String(v));
    if (s === '' || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return null;
    return Number(s);
  }

  /**
   * rows: mảng object {tên cột raw: giá trị} (giống pd.read_excel(dtype=str)).
   * Trả về {rows: [object theo OUT_COLS], warnings: [..]} hoặc ném lỗi nếu thiếu cột.
   */
  function convert(rawRows, mapping) {
    mapping = mapping || {};
    var ROSTER = ROSTER0.slice();
    (mapping.roster || []).forEach(function (r) { if (ROSTER.indexOf(r) < 0) ROSTER.push(r); });
    var UMAP = Object.assign({}, UMAP0, mapping.users || {});
    var VMAP = Object.assign({}, VMAP0);
    Object.keys(mapping.vendors || {}).forEach(function (k) { VMAP[k.toLowerCase()] = mapping.vendors[k]; });
    var AGENT_EXTRA = mapping.agents || {};

    var cols = {};
    rawRows.forEach(function (r) { Object.keys(r).forEach(function (k) { cols[k] = 1; }); });
    var missing = REQUIRED.filter(function (c) { return !cols[c]; });
    if (missing.length) { var e = new Error('File raw thiếu cột bắt buộc: ' + missing.join(', ')); e.missing = missing; throw e; }
    var df = rawRows.map(function (r) {
      var o = {};
      Object.keys(r).forEach(function (k) { var v = r[k]; o[k] = (v === '' || isNA(v)) ? null : v; });
      OPTIONAL.forEach(function (c) { if (!(c in o)) o[c] = null; });
      return o;
    });
    var warn = [];

    // ---- customer
    df.forEach(function (r) {
      var kt = custText(r['Tên khách hàng']), kn = custName(r['Tên']);
      var p = pick(kt, kn);
      r.KH = p[0]; r.KH_nguon = p[1];
    });

    // ---- vendor
    df.forEach(function (r) {
      var v = pyStrip(String(r['Tên Hãng'] == null ? '' : r['Tên Hãng']));
      if (v === 'Sản phẩm khác') {
        var n = String(r['Tên'] || '').toLowerCase();
        r.Hang = n.indexOf('synology') >= 0 ? 'Synology' : n.indexOf('yokogawa') >= 0 ? 'Yokogawa' : 'Khác';
      } else r.Hang = v === '' ? null : (VMAP.hasOwnProperty(v.toLowerCase()) ? VMAP[v.toLowerCase()] : v);
    });

    // ---- agent
    function akey(s) { return collapse(pyStrip(nfc(s).toLowerCase())); }
    var AGENT_ALIAS = { 'wintech': 'Wintech', 'công ty cổ phần wintech': 'Wintech', 'công ty cổ phần phát triển đô thị wintech': 'Wintech', 'ntshn': 'NTSHN', 'nts hn': 'NTSHN' };
    Object.keys(AGENT_EXTRA).forEach(function (k) { AGENT_ALIAS[akey(k)] = AGENT_EXTRA[k]; });
    function agent(s) {
      if (!isStr(s) || !pyStrip(s)) return null;
      var k = akey(s);
      if (AGENT_ALIAS.hasOwnProperty(k)) return AGENT_ALIAS[k];
      return pyStrip(collapse(s.replace(/ /g, ' ')));
    }
    df.forEach(function (r) { r.Dai_ly = agent(isNA(r['Đại lý']) ? null : String(r['Đại lý'])); });
    var canon = {}, seen = [];
    df.forEach(function (r) { var v = r.Dai_ly; if (v == null) return; var k = akey(v); if (!(k in canon)) canon[k] = v; if (seen.indexOf(v) < 0) seen.push(v); });
    seen.forEach(function (v) { var k = akey(v); if (pyIsUpper(canon[k]) && !pyIsUpper(v)) canon[k] = v; });
    df.forEach(function (r) { if (r.Dai_ly != null) r.Dai_ly = canon[akey(r.Dai_ly)]; });

    // ---- people
    function user(s) { return isStr(s) ? pyStrip(s.replace(/^@+/, '')) : (isNA(s) ? null : pyStrip(String(s).replace(/^@+/, ''))); }
    var unm = {};
    df.forEach(function (r) {
      r.Nguoi_tao = user(r['Người tạo']); r.Nguoi_thuc_hien = user(r['Người thực hiện']);
      r.Nguoi_tao_ten = r.Nguoi_tao == null ? null : (UMAP.hasOwnProperty(r.Nguoi_tao) ? UMAP[r.Nguoi_tao] : r.Nguoi_tao);
      r.Nguoi_thuc_hien_ten = r.Nguoi_thuc_hien == null ? null : (UMAP.hasOwnProperty(r.Nguoi_thuc_hien) ? UMAP[r.Nguoi_thuc_hien] : r.Nguoi_thuc_hien);
      [r.Nguoi_tao, r.Nguoi_thuc_hien].forEach(function (u) { if (u != null && !UMAP.hasOwnProperty(u)) unm[u] = 1; });
    });
    function nk(x) { return pyStrip(collapse(x.toLowerCase().replace(/đ/g, 'd').normalize('NFD').replace(/\p{Mn}/gu, ''))); }
    var RK = {}; ROSTER.forEach(function (r) { RK[nk(r)] = r; });
    function salefix(x) {
      if (isNA(x)) return null;
      x = pyStrip(collapse(String(x)));
      if (nfc(x.toLowerCase()) === nfc('nguyễn hồng nhung')) return 'Nguyễn Thị Hồng Nhung';
      var k = nk(x); return RK.hasOwnProperty(k) ? RK[k] : x;
    }
    df.forEach(function (r) { r.Sale_c = salefix(r['Sale']); r.Sale_trong_ds = ROSTER.indexOf(r.Sale_c) >= 0; });

    // ---- dates
    var badDate = {};
    df.forEach(function (r) {
      r.d_tao = parseDT(r['Tạo lúc'], true); r.d_cn = parseDT(r['Cập nhật lần cuối lúc'], true);
      r.d_giao = parseDT(r['Thời gian giao hàng'], false); r.d_ls = parseDT(r['License start date'], false); r.d_le = parseDT(r['License end date'], false);
      r.d_sla = parseDT(r['SLA/Thời hạn'], true);
      var done = r['Trạng thái'] === 'Hoàn thành';
      r.Thang_tao = r.d_tao == null ? null : ymOf(r.d_tao);
      r.TG_xu_ly = (done && r.d_tao != null && r.d_cn != null) ? npRound1(((r.d_cn - r.d_tao) * 1e6) * 1e-9 / 86400) : null;
      if (r['Tạo lúc'] != null && r.d_tao == null) badDate['Tạo lúc'] = 1;
      if (r['Thời gian giao hàng'] != null && r.d_giao == null) badDate['Thời gian giao hàng'] = 1;
      if (r['License end date'] != null && r.d_le == null) badDate['License end date'] = 1;
    });

    // ---- text
    df.forEach(function (r) {
      r.Mo_ta = mask(r['Mô tả phiếu']); r.BOM = mask(r['Thông tin hàng hóa (BOM)']);
      r.Ten_phieu = isNA(r['Tên']) ? null : pyStrip(collapse(String(r['Tên'])));
    });

    // ---- flags
    df.forEach(function (r) {
      var f = [], st = r['Trạng thái'], hd = r['Tình trạng Hợp đồng với khách hàng'];
      if (st === 'Hoàn thành' && hd === 'Chưa ký') f.push('Hoàn thành nhưng HĐ chưa ký');
      if (st === 'Hoàn thành' && hd === 'KH đã ký nhưng chưa trả Hợp đồng') f.push('Hoàn thành nhưng KH chưa trả HĐ');
      if (r['Quá hạn'] === 'Có') f.push('Quá hạn SLA');
      if (r.KH_nguon !== 'Form khách hàng') f.push('Tên KH lấy từ tên phiếu/không rõ - cần review');
      if (isNA(r['Đại lý'])) f.push('Thiếu đại lý');
      r.Co_review = f.length ? f.join('; ') : null;
    });

    // ---- output
    var out = df.map(function (r, i) {
      var no = toNumber(r['No.']);
      var id = toNumber(r['ID']);
      var o = {};
      o['No. (gốc)'] = no == null ? i + 1 : Math.trunc(no);
      o['ID phiếu'] = id == null ? null : id;
      o['Tên phiếu'] = r.Ten_phieu; o['Hãng'] = r.Hang; o['Khách hàng (chuẩn hóa)'] = r.KH; o['Nguồn tên KH'] = r.KH_nguon;
      o['Đại lý'] = r.Dai_ly; o['Sale'] = r.Sale_c; o['Người tạo'] = r.Nguoi_tao_ten; o['Người thực hiện'] = r.Nguoi_thuc_hien_ten;
      o['Luồng nhập hàng'] = r['Luồng nhập hàng']; o['Tình trạng HĐ với KH'] = r['Tình trạng Hợp đồng với khách hàng'];
      o['Công ty ký HĐ (*)'] = r['Công ty ký HĐ']; o['Biên bản nghiệm thu/thanh lý (*)'] = r['Biên bản nghiệm thu thanh lý'];
      o['Ngày tạo'] = r.d_tao; o['Cập nhật lần cuối'] = r.d_cn; o['Tháng tạo'] = r.Thang_tao; o['Ngày giao hàng'] = r.d_giao;
      o['License start (*)'] = r.d_ls; o['License end'] = r.d_le; o['Bước hiện tại'] = r['Tên khối']; o['Trạng thái'] = r['Trạng thái'];
      o['Thời gian xử lý (ngày, phiếu hoàn thành)'] = r.TG_xu_ly; o['SLA/Hạn (*)'] = r.d_sla; o['Quá hạn (*)'] = r['Quá hạn'];
      o['Cờ review'] = r.Co_review; o['Mô tả phiếu'] = r.Mo_ta; o['BOM'] = r.BOM;
      o['Người tạo (username)'] = r.Nguoi_tao_ten; o['Người thực hiện (username)'] = r.Nguoi_thuc_hien_ten;
      return o;
    });

    // ---- cảnh báo
    var um = Object.keys(unm).sort();
    if (um.length) warn.push('Username chưa có trong mapping (giữ nguyên username): ' + um.join(', '));
    var vset = {}; Object.keys(VMAP).forEach(function (k) { vset[VMAP[k]] = 1; }); vset.Synology = vset.Yokogawa = vset['Khác'] = 1;
    var badv = Array.from(new Set(df.map(function (r) { return r.Hang; }).filter(function (v) { return v != null && !vset[v]; }))).sort();
    if (badv.length) warn.push('Hãng chưa có trong bảng chuẩn hóa (giữ nguyên): ' + badv.join(', '));
    var ns = Array.from(new Set(df.filter(function (r) { return !r.Sale_trong_ds && r.Sale_c != null; }).map(function (r) { return r.Sale_c; }))).sort();
    if (ns.length) warn.push('Sale không có trong danh sách nhân sự: ' + ns.join(', '));
    var bd = ['Tạo lúc', 'Thời gian giao hàng', 'License end date'].filter(function (c) { return badDate[c]; });
    if (bd.length) warn.push('Có ngày sai định dạng ở cột: ' + bd.join(', '));
    return { rows: out, warnings: warn };
  }

  // ---------------- đọc/ghi Excel (cần thư viện XLSX = xlsx-js-style) ----------------
  function cellToPy(v) {
    // Mô phỏng pd.read_excel(dtype=str): số → chuỗi, ngày → giữ Date để parse linh hoạt
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return v;
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(v);
    if (typeof v === 'boolean') return v ? 'True' : 'False';
    return String(v);
  }
  // Đọc workbook với cellDates:false; ô ngày (số seri có định dạng ngày) → {__serial} để tránh lệch múi giờ
  function isDateCell(XLSX, ws, r, c) {
    var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
    return !!(cell && cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z));
  }
  function sheetToRows(XLSX, ws) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) return { header: [], rows: [] };
    var r0 = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s : { r: 0, c: 0 };
    var header = aoa[0].map(function (h) { return h == null ? '' : String(h); });
    var rows = [];
    for (var i = 1; i < aoa.length; i++) {
      var a = aoa[i], o = {}, any = false;
      header.forEach(function (h, j) {
        if (!h) return;
        var v = (typeof a[j] === 'number' && isDateCell(XLSX, ws, r0.r + i, r0.c + j)) ? { __serial: a[j] } : cellToPy(a[j]);
        o[h] = v; if (v != null) any = true;
      });
      if (any) rows.push(o);
    }
    return { header: header, rows: rows };
  }
  function toSerial(t) { return t / 86400000 + 25569; } // ms naive → số seri Excel

  function writeClean(XLSX, rows) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, cleanSheet(XLSX, rows), SHEET_NAME);
    return freezeAndWrite(XLSX, wb, { [SHEET_NAME]: 'A2' });
  }
  // Sheet "Data_sạch" đúng định dạng write_xlsx() của script Python
  function cleanSheet(XLSX, rows) {
    var aoa = [OUT_COLS.slice()];
    rows.forEach(function (r) {
      aoa.push(OUT_COLS.map(function (c) {
        var v = r[c];
        if (v == null) return null;
        if (DATETIME_COLS.indexOf(c) >= 0) return { t: 'n', v: toSerial(v), z: 'yyyy-mm-dd hh:mm' };
        if (DATE_COLS.indexOf(c) >= 0) return { t: 'n', v: toSerial(v), z: 'yyyy-mm-dd' };
        return v;
      }));
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var hs = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { patternType: 'solid', fgColor: { rgb: '1F4E78' } }, alignment: { wrapText: true, vertical: 'center' } };
    OUT_COLS.forEach(function (c, i) { var a = XLSX.utils.encode_cell({ r: 0, c: i }); ws[a].s = hs; });
    ws['!cols'] = COL_WIDTHS.map(function (w) { return { width: w }; });
    ws['!rows'] = [{ hpt: 45 }];
    ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }

  // Ghi workbook + chèn "freeze panes" (SheetJS bản cộng đồng không hỗ trợ sẵn)
  function freezeAndWrite(XLSX, wb, freeze) {
    var data = XLSX.write(wb, { type: 'array', bookType: 'xlsx', bookSST: true, compression: true });
    if (!freeze) return new Uint8Array(data);
    var cfb = XLSX.CFB.read(new Uint8Array(data), { type: 'array' });
    wb.SheetNames.forEach(function (name, i) {
      var cell = freeze[name]; if (!cell) return;
      var rc = XLSX.utils.decode_cell(cell);
      var idx = cfb.FullPaths.findIndex(function (p) { return p.endsWith('xl/worksheets/sheet' + (i + 1) + '.xml'); });
      if (idx < 0) return;
      var e = cfb.FileIndex[idx];
      var xml = new TextDecoder('utf-8').decode(e.content);
      var attrs = (rc.c ? ' xSplit="' + rc.c + '"' : '') + (rc.r ? ' ySplit="' + rc.r + '"' : '');
      var pane = rc.r && rc.c ? 'bottomRight' : rc.r ? 'bottomLeft' : 'topRight';
      xml = xml.replace(/<sheetView ([^>]*?)\/>/, '<sheetView $1><pane' + attrs + ' topLeftCell="' + cell + '" activePane="' + pane + '" state="frozen"/></sheetView>');
      e.content = new TextEncoder().encode(xml);
      e.size = e.content.length;
    });
    var out = XLSX.CFB.write(cfb, { type: 'array', fileType: 'zip', compression: true });
    return out instanceof Uint8Array ? out : new Uint8Array(out);
  }

  // Đọc file đã là template chuẩn (Data_sạch) → rows giống output convert
  function readClean(XLSX, ws) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    var header = (aoa[0] || []).map(function (h) { return h == null ? '' : String(h).trim(); });
    var rows = [];
    for (var i = 1; i < aoa.length; i++) {
      var a = aoa[i], o = {}, any = false;
      OUT_COLS.forEach(function (c) {
        var j = header.indexOf(c), v = j >= 0 ? a[j] : null;
        if (v === '' || v === undefined) v = null;
        if (v != null && (DATETIME_COLS.indexOf(c) >= 0 || DATE_COLS.indexOf(c) >= 0)) {
          if (typeof v === 'number') v = Math.round((v - 25569) * 1440) * 60000;
          else if (v instanceof Date) v = Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes());
          else { var s = String(v).trim(), m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/.exec(s); v = m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : (parseDT(s, true) || parseDT(s, false)); }
        }
        if (v != null && (c === 'No. (gốc)' || c === 'ID phiếu' || c === 'Thời gian xử lý (ngày, phiếu hoàn thành)')) { var n = toNumber(v); v = n; }
        if (v != null && c === 'Tháng tạo') v = String(v);
        if (v != null) any = true;
        o[c] = v;
      });
      if (any) rows.push(o);
    }
    return rows;
  }

  var api = {
    SHEET_NAME: SHEET_NAME, REQUIRED: REQUIRED, OPTIONAL: OPTIONAL, OUT_COLS: OUT_COLS, DATE_COLS: DATE_COLS, DATETIME_COLS: DATETIME_COLS,
    convert: convert, sheetToRows: sheetToRows, writeClean: writeClean, cleanSheet: cleanSheet, readClean: readClean, freezeAndWrite: freezeAndWrite,
    toSerial: toSerial, parseDT: parseDT, ymOf: ymOf
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.NTSConvert = api;
})(typeof window !== 'undefined' ? window : this);
