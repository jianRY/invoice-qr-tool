#!/usr/bin/env bash
#
# 发票二维码识别下载工具 —— 网站一键部署脚本（Linux + Nginx）
#
# 把本目录下的网站文件（index.html / download.html / assets / downloads）
# 拷贝到 Web 根目录，并可自动写好 Nginx 站点配置、重载 Nginx。
#
# 用法：
#   sudo ./deploy.sh                  # 默认装到 /var/www/invoice-qr-tool 并配置 nginx
#   sudo ./deploy.sh -d /srv/web      # 自定义 Web 根目录
#   sudo ./deploy.sh -h example.com   # 指定域名（写入 nginx server_name）
#   sudo ./deploy.sh --no-nginx       # 仅拷贝文件，不碰 nginx（适合已有 Web 服务器）
#
set -euo pipefail

# 脚本所在目录即网站根目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

WEB_ROOT="/var/www/invoice-qr-tool"
SERVER_NAME="_"
SETUP_NGINX=1

while [[ $# -gt 0 ]]; do
  case "$1" in
    -d|--dir)  WEB_ROOT="$2"; shift 2;;
    -h|--host) SERVER_NAME="$2"; shift 2;;
    --no-nginx) SETUP_NGINX=0; shift;;
    *) echo "未知参数: $1"; echo "用法: sudo ./deploy.sh [-d 目录] [-h 域名] [--no-nginx]"; exit 1;;
  esac
done

# 需要拷贝的网站内容（排除部署脚本与说明本身）
FILES=(index.html download.html assets downloads)

echo "==> 部署目录: $WEB_ROOT"
mkdir -p "$WEB_ROOT"

for f in "${FILES[@]}"; do
  if [[ -e "$SCRIPT_DIR/$f" ]]; then
    cp -r "$SCRIPT_DIR/$f" "$WEB_ROOT/"
    echo "    已拷贝: $f"
  else
    echo "    跳过（不存在）: $f"
  fi
done

# ---- Nginx 配置（可选）----
if [[ $SETUP_NGINX -eq 1 ]]; then
  if ! command -v nginx >/dev/null 2>&1; then
    echo "!! 未检测到 nginx，已跳过 nginx 配置。请自行把 Web 服务器根目录指向: $WEB_ROOT"
  else
    CONF="/etc/nginx/sites-available/invoice-qr-tool.conf"
    ENABLED="/etc/nginx/sites-enabled/invoice-qr-tool.conf"
    cat > "$CONF" <<EOF
server {
    listen 80;
    server_name ${SERVER_NAME};
    root ${WEB_ROOT};
    index index.html;

    location / {
        try_files \$uri \$uri/ =404;
    }

    # exe 安装包较大，允许断点续传并设缓存
    location /downloads/ {
        expires 1d;
        add_header Accept-Ranges bytes;
    }
}
EOF
    ln -sf "$CONF" "$ENABLED"
    echo "==> 已写入 nginx 配置: $CONF"
    if nginx -t; then
      systemctl reload nginx 2>/dev/null || service nginx reload 2>/dev/null || true
      echo "==> 已重载 nginx"
    else
      echo "!! nginx 配置测试失败，请检查后手动执行: nginx -t && systemctl reload nginx"
    fi
  fi
fi

echo ""
echo "✔ 部署完成！网站文件位于: $WEB_ROOT"
echo "  浏览器访问: http://${SERVER_NAME}/  （若指定了域名，请确保 DNS 已解析到本机）"
