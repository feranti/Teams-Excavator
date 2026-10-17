#!/usr/bin/env bash
# Renders every store_*.html in this folder to a 1280x800 PNG beside it
# (the 800x500 layout at 1.6x). Pass --scale N for another size.
cd "$(dirname "$0")"
SCALE=1.6
[ "$1" = "--scale" ] && SCALE="$2"
CHROME="/c/Program Files/Google/Chrome/Application/chrome.exe"
DIR=$(pwd -W 2>/dev/null || pwd)
W=$(awk "BEGIN{print 800*$SCALE}")
H=$(awk "BEGIN{print 500*$SCALE}")
for f in store_*.html; do
    out="${f%.html}_${W}x${H}.png"
    "$CHROME" --headless=new --disable-gpu --hide-scrollbars --allow-file-access-from-files \
        --force-device-scale-factor="$SCALE" --window-size=800,500 \
        --screenshot="$DIR/$out" "file:///$DIR/$f" 2>/dev/null
    echo "$out"
done
