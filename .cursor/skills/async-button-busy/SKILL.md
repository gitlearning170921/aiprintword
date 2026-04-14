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

## 文案配合

进行中文案与成功/失败提示分开：进行中用按钮内短句；结果用页面上的 `errMsg` / `signerErrMsg` / `batchResultMsg` 等区域。

## 配置沉淀（避免重复问题）

若某行为/策略需要可配置（例如 FTP 连接模式、失败回退策略），优先：

1. 在 `runtime_settings/registry.py` 注册为系统设置项（带分组/说明/是否 secret）。
2. 在代码中使用 `runtime_settings.resolve.get_setting()` 读取（避免到处查 `.env`）。
3. 在 `static/settings.html` 的分组顺序里加入对应 group（如 `ftp`），保证可在设置页管理与落库。
