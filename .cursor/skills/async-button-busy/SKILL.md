---
name: async-button-busy
description: 在 Web 前端为触发 fetch 或耗时逻辑的按钮提供统一的「处理中」禁用与可访问性状态；适用于本仓库 sign-page 及类似页面。
---

# 异步按钮忙碌态（项目约定）

## 何时使用

用户点击按钮后若需等待网络或异步任务，必须能立刻看出「已响应、在处理」，避免误以为未点到或卡死。

## 本仓库做法

在 `static/js/sign-page.js` 中使用 `withButtonBusy(btn, label, fn, opt)`：

1. 设置 `aria-busy="true"`、`disabled = true`、按钮内容为 `.spinner` + 简短文案（如「保存中…」）。
2. `fn` 返回 `Promise`；在 `finally` 中恢复 `innerHTML`，默认恢复 `disabled`。
3. 若结束后 `disabled` 必须由业务重算（例如依赖勾选状态），传入 `opt.skipRestoreDisabled: true`，并在链式 `finally` 里调用相应更新函数。

## 新增异步按钮时

- 校验逻辑仍在 `withButtonBusy` **之外**完成，避免校验失败时按钮已被禁用。
- 与 `updateBatchUi`、`updateSubmitState` 等状态函数协调，避免 `finally` 覆盖正确的 `disabled`。

## 文案与结果提示位置（防重复）

- **进行中**：仍用按钮内 `.spinner` + 短句（如「保存中…」）。
- **结果/错误**：写在**触发该操作的控件所属视觉组内**（可在按钮**右侧或换行后的正下方**，以易读、不与其他区块混淆为准），不要同一消息多处堆叠。
- **文件签名卡片**（`static/sign.html`）：与「生成已签名文档 / 批量签」相关的 **`errMsg`（错误）与 `batchResultMsg`（结果摘要）** 放在**「签名来源」下拉及高级选项块之后、同一卡片内**，与提交按钮**不同行**（避免误当成其它按钮的反馈）。其它按钮（刷新列表、重新识别、保存到列表、表内载入等）仍各自用 **`btn-inline-feedback`**，与触发控件**同一行或紧接其下一行**均可，以美观、易看见为原则。
- **签署人区**：用 `signerErrMsg` 等已有 `btn-inline-feedback`，与「添加 / 刷新」同卡片内紧邻。
- **脚本整页级失败**（如 `sign-page.js` 未加载）：保留 `signBootstrapBanner` 顶栏即可；**不要**再把同一段文字写进「签署人列表」「文件列表」的 hint 里（避免与 `errMsg` 三处重复）。`showScriptLoadFail` 仅更新 banner + `errMsg`。

## 配置沉淀（避免重复问题）

若某行为/策略需要可配置（例如 FTP 连接模式、失败回退策略），优先：

1. 在 `runtime_settings/registry.py` 注册为系统设置项（带分组/说明/是否 secret）。
2. 在代码中使用 `runtime_settings.resolve.get_setting()` 读取（避免到处查 `.env`）。
3. 在 `static/settings.html` 的分组顺序里加入对应 group（如 `ftp`），保证可在设置页管理与落库。
