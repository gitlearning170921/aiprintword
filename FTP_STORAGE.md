# FTP 文件存储约定

本项目采用 **FTP 存文件**、**MySQL 存记录** 的策略。

## 配置项（.env）

- `FTP_HOST` / `FTP_PORT` / `FTP_USER` / `FTP_PASSWORD`
- `FTP_BASE_DIR`：父目录（例如 `/upload`）
- `FTP_APP_DIR`：应用目录名（默认 `aiprintword`）

最终实际落盘根目录为：

\[
  \text{FTP_ROOT} = \text{FTP_BASE_DIR}/\text{FTP_APP_DIR}
\]

例如：`/upload/aiprintword`

## 目录规划（FTP_ROOT 下）

- **原始上传（批处理）**：`batch/upload/<tmp_dir>/<rel>`
- **批处理 ZIP 导出包**：`batch/exports/<token>.zip`
- **在线签名待签文件**：`sign/inbox/<file_id>/<original_name>`
- **在线签名已签名产物**：`sign/output/<signed_id>/<output_name>`
- **签署人可复用笔迹 PNG**：`sign/strokes/<signer_id>/sig.png`、`sign/strokes/<signer_id>/date.png`

## MySQL 只存什么

- **待签文件**：`sign_uploaded_file` 里存 `ftp_path/file_size/sha256` 与名称/扩展名等记录
- **已签名结果**：`sign_signed_output` 里存 `ftp_path/file_size/sha256` 与关联信息
- **签署人笔迹**：`sign_signer_stroke` 里存 `sig_ftp_path/date_ftp_path` 与校验信息
- **兼容旧数据**：历史遗留的 `file_data/sig_png/date_png`（BLOB）仍可读取；迁移后可置 NULL 释放空间

## 迁移（MySQL BLOB -> FTP）

脚本：

- `python migrate_sign_mysql_blobs_to_ftp.py`

管理接口（管理员 token）：

- `POST /api/admin/sign/migrate-mysql-blobs-to-ftp`

可选参数：

- `limit`：每次扫描上限
- `clear_blob`：迁移成功后是否清空 BLOB（默认 true）

