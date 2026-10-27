#!/bin/zsh
cd "${0:A:h}/网页实验" || exit 1
PORT=8766
URL="http://127.0.0.1:${PORT}/%E4%BA%91%E5%AE%9D%E4%BA%8C%E6%96%B9%E8%BF%9E%E7%BB%AD%E7%BA%B9%E6%A0%B7%E8%AE%BE%E8%AE%A1%E5%B7%A5%E5%9D%8A_v1.html"
open "$URL"
python3 -m http.server "$PORT" --bind 127.0.0.1
