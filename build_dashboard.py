# -*- coding: utf-8 -*-
"""
Builds Stock_Dashboard.html from all '*Stock statement.xlsx' files in this folder.

Monthly workflow:
  1. Drop the new month's 'XXX-YY Stock statement.xlsx' here (same sheet layout)
  2. Run:  python build_dashboard.py
  3. Open Stock_Dashboard.html in any browser
"""
import json
import os

import extract_stock_data

HERE = os.path.dirname(os.path.abspath(__file__))


def build(template_name, payload, logo_b64, outputs):
    with open(os.path.join(HERE, template_name), encoding='utf-8') as f:
        html = f.read()
    html = html.replace('/*__STOCK_DATA__*/', payload, 1)
    if logo_b64:
        html = html.replace('__LOGO_B64__', logo_b64)
    for name in outputs:
        out = os.path.join(HERE, name)
        with open(out, 'w', encoding='utf-8') as f:
            f.write(html)
        print('Built', out)


def main():
    months = extract_stock_data.main()
    payload = json.dumps({'months': months}, ensure_ascii=False)
    logo_b64 = ''
    logo_path = os.path.join(HERE, 'concorde_logo.png')
    if os.path.exists(logo_path):
        import base64
        with open(logo_path, 'rb') as f:
            logo_b64 = base64.b64encode(f.read()).decode()
    # Stock_Dashboard.html for local double-click use, index.html for web hosting (Vercel)
    build('dashboard_template.html', payload, logo_b64, ('Stock_Dashboard.html', 'index.html'))
    build('erp_template.html', payload, logo_b64, ('erp.html',))


if __name__ == '__main__':
    main()
