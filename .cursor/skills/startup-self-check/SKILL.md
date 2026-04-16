# Skill: 启动自检与“页面打不开”排障

适用场景：用户反馈「服务已启动但页面打不开 / 一直转圈 / HTTP 超时」，尤其是 Windows 环境。

## 目标

- 优先排除 **5050 端口被多个进程占用/旧进程未关闭** 导致的“连到旧进程/卡死进程”。
- 给出用户可直接复制执行的排查与关闭命令。
- 建议加入或确认 `app.py` 已包含启动时端口占用自检（单实例）。

## 操作清单（按优先级）

### 1) 先判断是否端口可连通

- 让用户执行（PowerShell）：
  - `Test-NetConnection 127.0.0.1 -Port 5050`

若 TCP 通但浏览器/HTTP 超时，极可能是旧进程/卡死进程或多个监听冲突。

### 2) 检测是否存在多个监听进程（Windows）

让用户执行：

- `netstat -ano | findstr ":5050"`

观察 `LISTENING` 行的 PID：

- 若出现 **多个 LISTENING PID**，说明同时起了多个服务进程。

再确认进程：

- `tasklist /FI "PID eq <PID>" /FO LIST`

关闭多余进程：

- `taskkill /PID <PID> /F`

### 3) 只保留一个进程后重新启动

在项目目录启动：

- `python app.py`

浏览器访问：

- `http://127.0.0.1:5050/sign`
- `http://127.0.0.1:5050/sign/materials`

并 `Ctrl+F5` 强刷。

## 代码沉淀要求

- `app.py` 在 `app.run(..., port=5050)` 之前必须做端口占用自检：
  - 若占用：打印“已有其它 app.py 在跑，先关掉它”并退出（非 0）。
  - 同时打印 Windows/Linux/macOS 的排查命令（见上）。

