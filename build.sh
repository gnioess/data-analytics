#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

mkdir -p dist docs

TEMPLATE="src/index_template.html"
VENDOR="vendor/xlsx.mini.min.js"
APP="src/app.js"

# Two build variants from the same source:
#   dist/ERP_Analytics.html — downloadable/local: manual Excel upload only
#     (Google's login can't run from a file:// origin, so Drive is hidden there)
#   docs/index.html         — online (GitHub Pages): Google Drive only, no
#     manual upload UI, since it always has a real https:// origin
python3 - "$TEMPLATE" "$VENDOR" "$APP" <<'PY'
import sys
template_path, vendor_path, app_path = sys.argv[1:4]
template = open(template_path, encoding='utf-8').read()
vendor = open(vendor_path, encoding='utf-8').read()
app = open(app_path, encoding='utf-8').read()
base = template.replace('<!--VENDOR_JS-->', '<script>\n' + vendor + '\n</script>')
base = base.replace('<!--APP_JS-->', '<script>\n' + app + '\n</script>')

targets = [
    ('dist/ERP_Analytics.html', 'local'),
    ('docs/index.html', 'online'),
]
for out_path, mode in targets:
    out = base.replace('BUILD_MODE_TOKEN', mode)
    open(out_path, 'w', encoding='utf-8').write(out)
    print(f"Wrote {out_path} ({len(out)} bytes) [mode={mode}]")
PY
