# 发票二维码识别下载工具 · 网站部署说明

本包是一个**纯静态网站**（无后端、无数据库、无外部依赖），用于展示软件功能与使用场景，
并提供软件本体的**本站直链下载**（exe 已随包自带，托管在 `downloads/` 目录）。

---

## 一、包内文件

```
invoice-qr-tool-website/
├── index.html                       # 首页：功能介绍 + 使用场景 + 三步使用
├── download.html                    # 下载页：下载按钮直链 downloads/ 下的 exe
├── assets/
│   ├── style.css                    # 样式（浅色主题、响应式）
│   └── app-icon.png                 # 图标 / favicon
├── downloads/
│   └── InvoiceQRDownloader_v3.6.exe # 软件本体（约 111 MB，本站托管）
├── deploy.sh                        # Linux + Nginx 一键部署脚本
└── DEPLOY.md                        # 本说明
```

> 因为是纯静态文件，任何 Web 服务器（Nginx / Apache / IIS / 对象存储静态托管）都能托管。

---

## 二、一键部署（推荐：Linux + Nginx）

1. 把整个包（保持目录结构）上传到服务器，例如 `/root/invoice-qr-tool-website/`。
2. 给脚本加执行权限并运行（需要 root，因为要写 `/var/www` 和 nginx 配置）：

   ```bash
   cd /root/invoice-qr-tool-website
   chmod +x deploy.sh
   sudo ./deploy.sh
   ```

   默认会把网站装到 `/var/www/invoice-qr-tool`，并自动写好 Nginx 站点、重载 Nginx。
   部署完成后访问 `http://服务器IP/` 即可。

### 常用自定义参数

| 命令 | 作用 |
| --- | --- |
| `sudo ./deploy.sh -d /srv/web` | 自定义 Web 根目录（不写到默认目录） |
| `sudo ./deploy.sh -h example.com` | 指定域名（写入 `server_name`，需自行配好 DNS） |
| `sudo ./deploy.sh --no-nginx` | 只拷贝文件、不碰 Nginx（适合你已有 Web 服务器） |

示例：绑定域名并部署到自定义目录

```bash
sudo ./deploy.sh -d /srv/invoice-qr-tool -h fapiao.example.com
```

---

## 三、其它服务器 / 已有 Web 服务器

### 已有 Nginx（只要文件）
用 `--no-nginx` 把文件拷到你想放的位置，然后自己加一段 `server { root ... }` 指向该目录即可。

### Apache
把包内容放到 `DocumentRoot` 下（如 `/var/www/html/invoice-qr-tool/`），
确保 `mod_rewrite` 或默认目录索引可用，`index.html` 会被自动作为首页。

### Windows / IIS
把包内容放到某个站点目录（如 `C:\inetpub\wwwroot\invoice-qr-tool\`），
在 IIS 管理器中新建站点或虚拟目录指向它，并确认 `.exe` 的 MIME 类型已被允许下载
（IIS 默认允许 `.exe` 下载；若被拦，在「MIME 类型」里添加 `.exe → application/octet-stream`）。

---

## 四、以后升级软件版本

1. 把新的 exe 改名成 `InvoiceQRDownloader_vX.Y.exe` 放进 `downloads/`。
2. 修改 `download.html` 里的两处下载链接（主按钮 + GitHub 备用镜像）以及版本徽章文字。
3. 重新运行 `sudo ./deploy.sh` 覆盖部署即可（脚本只会覆盖网站文件，不影响其它）。

---

## 五、关于「自动更新」

软件内置的「检查更新」仍通过 GitHub 公开 Release 判断最新版本，发现新版本会
自动下载安装、并把旧版本移入回收站。这与本网站自托管下载互不冲突——
用户从哪里下载到 exe 都不影响软件内部的自动升级机制。

---

## 六、安全与隐私提示

- 网站本身不收集任何信息；软件所有识别 / 下载 / 汇总都在用户本机完成，票据数据不出本机。
- `downloads/` 下的 exe 是普通二进制，建议保持目录仅可读、不可执行（防止被当作 CGI 执行）。
- 如需 HTTPS，可在 Nginx 前置一层证书（如 Let's Encrypt / certbot），网站无需改动。
