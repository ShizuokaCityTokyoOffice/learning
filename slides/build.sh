#!/bin/sh
# slides.html を PDF に変換する（Chromium ヘッドレス）
cd "$(dirname "$0")"
CHROME=$(ls /opt/pw-browsers/chromium-*/chrome-linux/chrome | head -1)
"$CHROME" --headless --no-sandbox --disable-gpu --no-pdf-header-footer \
  --print-to-pdf=slides.pdf "file://$PWD/slides.html"
