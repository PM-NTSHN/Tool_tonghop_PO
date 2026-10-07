#!/usr/bin/env python3
"""Đóng gói src/* + thư viện + logo thành 1 file HTML offline duy nhất: "Tool tổng hợp PO.html"."""
import base64, pathlib, sys
ROOT = pathlib.Path(__file__).resolve().parent
SRC, VENDOR = ROOT / 'src', ROOT / 'vendor'

def js(path):
    t = path.read_text(encoding='utf-8')
    return t.replace('</script', '<\\/script')  # tránh đóng thẻ script sớm

def main():
    html = (SRC / 'app.html').read_text(encoding='utf-8')
    logo = 'data:image/png;base64,' + base64.b64encode((SRC / 'logo_nts.png').read_bytes()).decode()
    parts = {
        '/*__CSS__*/': (SRC / 'app.css').read_text(encoding='utf-8'),
        '/*__XLSX__*/': js(VENDOR / 'xlsx-js-style.min.js'),
        '/*__CHART__*/': js(VENDOR / 'chart.umd.min.js'),
        '/*__CONVERTER__*/': js(SRC / 'converter.js'),
        '/*__APP__*/': js(SRC / 'app.js'),
    }
    for k, v in parts.items():
        assert html.count(k) == 1, k
        html = html.replace(k, v)
    html = html.replace('__LOGO__', logo)
    names = [sys.argv[1]] if len(sys.argv) > 1 else ['Tool tổng hợp PO.html', 'index.html']  # index.html cho GitHub Pages
    for n in names:
        out = ROOT / n
        out.write_text(html, encoding='utf-8')
        print(f'OK -> {out.name} ({out.stat().st_size/1024:.0f} KB)')

if __name__ == '__main__':
    main()
