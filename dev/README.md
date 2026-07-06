# aiprintword — 开发机脚本与目录说明

```
aiprintword/
├── app.py sign_handlers/ static/   ★ 签批/批量打印 产品代码
├── scripts/ docs/                  ★ 工具与文档
├── dev/
│   ├── git-no_tag/                 ◆ 日常 commit+push（不打 tag）
│   ├── git-tag_release/            ◆ 说明：不在主 Docker 发版链
│   └── local-run/                  ◆ 本机启动 5050 服务
└── data/                           ✗ 运行时数据
```

## 常用命令

```cmd
:: 日常提交
dev\git-no_tag\commit_push.bat "签批页优化"

:: 本机启动（Windows 服务器 / 开发联调）
dev\local-run\start_server.bat
```

aiword 页面3 通过 `AIPRINTWORD_BASE_URL` 指向本服务（默认 `http://<Windows IP>:5050`）。

主发版流程见 `..\aiword\dev\README.md`（aiprintword 需单独部署到 Windows，不打进 aiword Docker 包）。
