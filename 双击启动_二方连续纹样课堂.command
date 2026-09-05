#!/bin/zsh

WEB_DIR="/Users/chengqp/Desktop/tan"
WEB_PORT="8768"
WEB_PAGE="%E4%BA%8C%E6%96%B9%E8%BF%9E%E7%BB%AD%E7%BA%B9%E6%A0%B7_%E7%94%9F%E6%88%90%E8%A7%84%E5%88%99%E6%89%A7%E8%A1%8C%E4%B8%8E%E6%A3%80%E9%AA%8C_v2.html"
WEB_LOG="/tmp/frieze_rule_lab_8768.log"
WEB_URL="http://127.0.0.1:$WEB_PORT/$WEB_PAGE"

cd "$WEB_DIR" || exit 1

python3 -m http.server "$WEB_PORT" --bind 127.0.0.1 >"$WEB_LOG" 2>&1 &
WEB_SERVER_PID=$!

finish_web_server() {
  kill "$WEB_SERVER_PID" >/dev/null 2>&1
}
trap finish_web_server EXIT INT TERM

for WEB_TRY in 1 2 3 4 5; do
  if curl -fsS "$WEB_URL" >/dev/null 2>&1; then
    open "$WEB_URL"
    echo ""
    echo "课堂网页已经打开。"
    echo "上课期间请保留此窗口；结束后回到这里按回车键关闭服务。"
    read -r "?按回车键结束……"
    exit 0
  fi
  sleep 1
done

echo "网页服务未能启动，请把下面的日志发给技术支持："
echo "$WEB_LOG"
read -r "?按回车键关闭窗口……"
exit 1
