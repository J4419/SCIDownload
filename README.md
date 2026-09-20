# SCIDownload

> 当前发布版：**v1.6.1**　／　[使用说明.md](使用说明.md) · [SECURITY.md](SECURITY.md)

SCIDownload 用本机 Chrome、Edge 或 Chromium 打开文章页，并把当前用户**已有权访问**的正文 PDF 与可选补充材料（SI）保存到本地。它不提供机构订阅、不绕过付费墙，也不能让无权限的文章变成可下载。

主战场是 **Elsevier / ScienceDirect**（`10.1016/*`，含 JIA、Pedosphere）；同时已实测支持
**Wiley**（`10.1002` `10.1111` `10.2134` `10.2136`）、
**加拿大科学出版社**（`10.1139` `10.4141`）、
**CSIRO**（`10.1071`）、
**Copernicus**（`10.5194`）、
**PeerJ**（`10.7717`）、**Nature Portfolio / Scientific Reports**（`10.1038`）、
**MDPI**（`10.3390`）等出版商。
清单里混着这些 DOI 也能一个命令跑完。v1.5.2 起，对没有专门规则的非 ScienceDirect 页面，
程序还会尝试标准 `citation_pdf_url` 和页面中明确的 PDF 链接；v1.6.0 又增加了对浏览器原生 attachment 下载的捕获，因此像 MDPI 这类“手动点击会直接下载文件、不会在标签页打开 PDF”的页面也能被程序正确接管。只有专门规则与通用候选都失败后才记为 `unsupported_publisher`。详见 [使用说明.md](使用说明.md) 第 9 节。

项目只依赖 Python 标准库，支持 Windows、macOS 和 Linux。发布包内含标题反查 DOI、出口网络诊断和缓存清理工具。

## 使用前提

- Python 3.10 或更高版本，且 `python` 命令可用（**如果你的 `python` 不可用，见 [使用说明.md](使用说明.md) 开头「确认 python 命令能用」一节**）。
- Chrome、Edge 或 Chromium。
- 对目标内容有合法访问权，例如开放获取、校园网机构订阅，或已获授权的 CARSI、学校 VPN、EZproxy/机构登录会话。
- 遵守出版商协议、学校图书馆政策和适用法律。不要并发运行，不要把间隔调得不合理地小，也不要下载超出真实研究需要的内容。

## 三篇试跑

准备 `dois.txt`，每行一个 DOI：

```text
10.1016/j.example.2025.000001
10.1016/j.example.2025.000002
```

在解压后的目录运行。若当前目录只有一个 `.txt/.csv/.tsv/.xlsx/.xlsm` 清单，可以不写文件名：

```powershell
python 1.check_exit_ip.py
python 3.SCIDownload.py --out ./papers --limit 3 --si
```

程序会显示 `未指定清单文件；自动发现唯一候选：...`。如果当前目录有多个候选清单，程序会把它们列出来并停止，不会自行猜测；这时显式指定即可：

```powershell
python 3.SCIDownload.py dois.txt --out ./papers --limit 3 --si
```

`1.check_exit_ip.py` 只是网络排障工具，默认会**脱敏公网 IP**，并且只通过 HTTPS 查询公开出口信息。代理和分流可能让 Python 探测到的出口与浏览器实际访问链路不同；最终仍以文章页的真实授权状态为准。确需查看完整 IP 时才使用 `--show-full-ip`。**主程序默认不联系第三方 IP 查询服务**；只有显式加 `--ip-check` 才运行这一诊断。

若某篇在首次运行时返回 `no_pdf_link`（页面确认可访问但未解析到正文入口），**原样重跑同一条命令**即可 —— 这通常是浏览器首次会话尚未预热所致，不是权限问题。

三篇都正常后，去掉 `--limit 3`：

```powershell
python 3.SCIDownload.py dois.txt --out ./papers --si
```

首次需要机构登录时可加 `--login-wait 300`，在弹出的浏览器中完成登录。专用 profile 会保留会话。共享电脑上不要使用 `--keep-browser`，也不要分享 profile 目录。

## 输入和输出

输入支持 `.txt`、`.csv`、`.tsv`、`.xlsx` 和 `.xlsm`。**未填写输入文件名时，程序会扫描当前工作目录：唯一候选自动使用，多个候选则要求明确指定。** 表格可包含 `DOI`、`Num`、`Title` 列。DOI 会统一为小写并按首次出现稳定去重。只有标题时先运行：

```powershell
python 2.FindDOI.py titles.xlsx --out titles_DOI.csv --mailto you@example.edu
```

FindDOI 默认遵循系统代理；只有需要强制直连公开元数据 API 时才加 `--direct`。`--mailto` 是可选项：填写后邮箱会随请求发送给 OpenAlex/Crossref；不希望提供即可省略。自动填 DOI 的默认标题相似度门槛为 **0.97**，0.80～0.97 的边界结果留空供人工确认。

输出目录包含：

- `PDFs/`：正文 PDF；写入前验证文件头与长度，并使用原子替换。浏览器 attachment 下载也会先落到程序专用临时目录，不会使用系统默认 Downloads。
- `SupportingInformation/`：使用 `--si` 时保存的补充材料。
- `manifest.jsonl`：每次尝试一行；同一 DOI 的最后一条有效 JSON 记录代表最新状态。
- `run.log`：运行日志。

原命令重跑即可续跑。程序会验证磁盘上的正文 PDF，且对 manifest 标记为 `si_status=downloaded` 的补充材料也会逐个检查实际文件：正文/SI 完整会跳过，缺失或损坏时只补需要的部分。`--force` 会忽略断点状态。对尚未实现可靠 SI 枚举的出版商，状态会记为 `not_supported`，不会误写成“该文没有 SI”。

完整选项、逐参数说明、状态表、排错与暂停/恢复方法见 **[使用说明.md](使用说明.md)**。

## API key 与 `insttoken`

本项目保留的是浏览器工作流，**不需要 Elsevier API key**。若另行使用 Elsevier 官方 API：请求需要 API key；机构订阅有时可通过机构网络自动识别，校外或特定集成可能需要由机构/Elsevier 配置 `insttoken`。`insttoken` 不是个人在网页上随手生成的万能凭据，也不能替代学校没有购买的订阅。没有开放获取或相应订阅/API 权益时，付费全文依然不可用。

官方说明：[API authentication](https://dev.elsevier.com/tecdoc_api_authentication.html)。

## 隐私与安全

- 浏览器 profile 可能含 Cookie、机构登录态和访问记录，不要提交到 GitHub、打包或发给他人。
- 主程序默认对出口 IP、本地绝对路径和 URL 查询参数做脱敏；但日志、manifest、终端截图仍可能包含 DOI、出版商/机构相关信息和研究兴趣，公开前仍要检查。
- 代理环境变量可能含账号密码。诊断脚本只显示脱敏后的代理端点；公网 IP 默认也会脱敏。
- 远程调试只监听 `127.0.0.1`，且启动前会拒绝已占用端口。
- `4.ClearCache.py` 默认只清缓存；`--all` 会删除整个专用 profile，使用自定义 profile 时还要显式加 `--confirm-custom-profile`。

安全问题与敏感信息处理见 [SECURITY.md](SECURITY.md)。

## 许可证

[MIT](LICENSE)。许可证仅覆盖代码，不授予第三方内容、机构订阅或出版商服务的使用权。
