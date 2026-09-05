#!/bin/zsh

LAB_DIR="/Users/chengqp/Desktop/tan"
LAB_PORT="8770"
LAB_PAGE="%E4%BA%8C%E6%96%B9%E8%BF%9E%E7%BB%AD%E7%BA%B9%E6%A0%B7_%E5%AD%A6%E7%94%9F%E6%95%B0%E5%AD%A6%E7%A0%94%E7%A9%B6%E5%AE%9E%E9%AA%8C%E5%AE%A4_v4.html"
LAB_URL="http://127.0.0.1:$LAB_PORT/$LAB_PAGE"
LAB_LOG="/tmp/frieze_math_lab_v4_8770.log"

cd "$LAB_DIR" || exit 1
python3 -m http.server "$LAB_PORT" --bind 127.0.0.1 >"$LAB_LOG" 2>&1 &
LAB_SERVER_PID=$!

close_lab_server() {
  kill "$LAB_SERVER_PID" >/dev/null 2>&1
}
trap close_lab_server EXIT INT TERM

for LAB_TRY in 1 2 3 4 5; do
  if curl -fsS "$LAB_URL" >/dev/null 2>&1; then
    open "$LAB_URL"
    echo ""
    echo "连续纹样数学研究室 v4 已打开。"
    echo "课堂使用期间请保留此窗口；下课后按回车键关闭。"
    read -r "?按回车键结束……"
    exit 0
  fi
  sleep 1
done

echo "启动失败。日志：$LAB_LOG"
read -r "?按回车键关闭……"
exit 1
