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


def main():
    months = extract_stock_data.main()
    with open(os.path.join(HERE, 'dashboard_template.html'), encoding='utf-8') as f:
        template = f.read()
    payload = json.dumps({'months': months}, ensure_ascii=False)
    html = template.replace('/*__STOCK_DATA__*/', payload, 1)
    # Stock_Dashboard.html for local double-click use, index.html for web hosting (Vercel)
    for name in ('Stock_Dashboard.html', 'index.html'):
        out = os.path.join(HERE, name)
        with open(out, 'w', encoding='utf-8') as f:
            f.write(html)
        print('Built', out)


if __name__ == '__main__':
    main()
