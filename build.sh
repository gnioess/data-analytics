#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

OUT="dist/Panel_INDAMA.html"
mkdir -p dist

TEMPLATE="src/index_template.html"
VENDOR="vendor/xlsx.full.min.js"
APP="src/app.js"

python3 - "$TEMPLATE" "$VENDOR" "$APP" "$OUT" <<'PY'
import sys
template_path, vendor_path, app_path, out_path = sys.argv[1:5]
template = open(template_path, encoding='utf-8').read()
vendor = open(vendor_path, encoding='utf-8').read()
app = open(app_path, encoding='utf-8').read()
out = template.replace('<!--VENDOR_JS-->', '<script>\n' + vendor + '\n</script>')
out = out.replace('<!--APP_JS-->', '<script>\n' + app + '\n</script>')
open(out_path, 'w', encoding='utf-8').write(out)
print(f"Wrote {out_path} ({len(out)} bytes)")
PY
