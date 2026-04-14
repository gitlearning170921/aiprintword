# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional


@dataclass(frozen=True)
class SettingMeta:
    key: str
    default: Any
    value_type: str  # bool | int | float | str
    group: str
    label: str
    is_secret: bool = False
    description: str = ""


def _b(default: bool) -> bool:
    return default


REGISTRY: Dict[str, SettingMeta] = {}
_ORDER: List[str] = []


def _reg(meta: SettingMeta) -> None:
    REGISTRY[meta.key] = meta
    _ORDER.append(meta.key)


_reg(
    SettingMeta(
        "WORD_CONTENT_PRESERVE",
        True,
        "bool",
        "word",
        "Word 内容保真（减少激进清理）",
        description="1=开启，避免删水印图形、删删除线等高风险步骤",
    )
)
_reg(
    SettingMeta(
        "WORD_IMAGE_RISK_GUARD",
        True,
        "bool",
        "word",
        "Word 图片完整性风险检测",
    )
)
_reg(
    SettingMeta(
        "WORD_PRESERVE_LINKED_IMAGES",
        False,
        "bool",
        "word",
        "Word 链接图保全模式",
        description="先断链内嵌，选择性域更新",
    )
)
_reg(
    SettingMeta(
        "WORD_HEADER_FOOTER_LAYOUT_FIX",
        True,
        "bool",
        "word",
        "Word 页眉页脚自动修补",
    )
)
_reg(
    SettingMeta(
        "WORD_STEP_TIMEOUT_SEC",
        3600.0,
        "float",
        "word",
        "Word 单步超时（秒）",
        description="接受修订深扫等步骤的超时上限",
    )
)
_reg(
    SettingMeta(
        "WORD_SKIP_FILE_ON_TIMEOUT",
        True,
        "bool",
        "word",
        "Word 超时跳过文件",
    )
)
_reg(
    SettingMeta(
        "WORD_PRESERVE_PAGE_COUNT",
        True,
        "bool",
        "word",
        "Word 总页数保护",
        description="处理后页数变化则回滚备份",
    )
)
_reg(
    SettingMeta(
        "WORD_MANY_TABLE_ROWS_LITE_THRESHOLD",
        100,
        "int",
        "word",
        "表格行轻量模式阈值",
        description="docx 主文档内表格行（XML 中 w:tbl//w:tr，含嵌套表）合计≥此值走轻量路径；行数越多通常越耗时，调低更易进轻量",
    )
)
_reg(
    SettingMeta(
        "AIPRINTWORD_WORD_BACKUP_TEMP_DIR",
        "",
        "str",
        "word",
        "Word 页数保护备份目录",
        description="留空用系统 TEMP；设为本地 SSD 目录可减轻慢盘与杀毒实时扫描对备份副本的影响",
    )
)
_reg(
    SettingMeta(
        "USE_OFFICE",
        "",
        "str",
        "office",
        "使用 Microsoft Office",
        description="填 1/true/yes 使用 Word/Excel；留空则使用 WPS",
    )
)
_reg(
    SettingMeta(
        "AIPRINTWORD_LOG_LEVEL",
        "INFO",
        "str",
        "app",
        "日志级别",
        description="修改后需重启进程生效",
    )
)
_reg(
    SettingMeta(
        "AIPRINTWORD_ALLOWED_OUTPUT_PARENT",
        "",
        "str",
        "app",
        "允许即时保存的根目录",
        description="非空则 incremental_output_dir 必须在其下；留空不限制",
    )
)
_reg(
    SettingMeta(
        "AIPRINTWORD_HISTORY_MAX",
        50,
        "int",
        "app",
        "批量历史保留条数上限",
        description="范围 5–500",
    )
)
_reg(
    SettingMeta(
        "AIPRINTWORD_SSE_HEARTBEAT_SEC",
        120.0,
        "float",
        "app",
        "SSE 心跳间隔（秒）",
        description="范围 10–600",
    )
)
_reg(
    SettingMeta(
        "SIGN_MYSQL_MAX_FILES",
        500,
        "int",
        "sign",
        "签名待签文件数量上限",
        description="范围 1–10000",
    )
)
_reg(
    SettingMeta(
        "SIGN_MYSQL_MAX_SIGNED",
        2000,
        "int",
        "sign",
        "签名已签输出条数上限",
        description="范围 1–50000",
    )
)
_reg(
    SettingMeta(
        "SIGN_FTP_REQUIRED",
        False,
        "bool",
        "sign",
        "签名：强制上传 FTP 成功",
        description="1=生成/保存素材时必须上传 FTP 成功，否则返回失败；0=FTP 失败则回退 MySQL BLOB（仍可下载），并在相关记录中保存 ftp_last_error 便于排查",
    )
)
_reg(
    SettingMeta(
        "MYSQL_HOST",
        "",
        "str",
        "mysql",
        "MySQL 主机",
        is_secret=False,
    )
)
_reg(
    SettingMeta(
        "MYSQL_PORT",
        "3306",
        "str",
        "mysql",
        "MySQL 端口",
    )
)
_reg(
    SettingMeta(
        "MYSQL_DATABASE",
        "aiprintword_sign",
        "str",
        "mysql",
        "MySQL 数据库名",
    )
)
_reg(
    SettingMeta(
        "MYSQL_USER",
        "root",
        "str",
        "mysql",
        "MySQL 用户",
    )
)
_reg(
    SettingMeta(
        "MYSQL_PASSWORD",
        "",
        "str",
        "mysql",
        "MySQL 密码",
        is_secret=True,
    )
)
_reg(
    SettingMeta(
        "MYSQL_CHARSET",
        "utf8mb4",
        "str",
        "mysql",
        "MySQL 字符集",
    )
)

# ----------------
# FTP（签名素材/已签输出/待签文件等可上传到 FTP）
# ----------------
_reg(
    SettingMeta(
        "FTP_HOST",
        "10.26.1.221",
        "str",
        "ftp",
        "FTP 主机",
    )
)
_reg(
    SettingMeta(
        "FTP_PORT",
        "2121",
        "str",
        "ftp",
        "FTP 端口",
    )
)
_reg(
    SettingMeta(
        "FTP_USER",
        "aiwordftpuser",
        "str",
        "ftp",
        "FTP 用户",
    )
)
_reg(
    SettingMeta(
        "FTP_PASSWORD",
        "",
        "str",
        "ftp",
        "FTP 密码",
        is_secret=True,
    )
)
_reg(
    SettingMeta(
        "FTP_BASE_DIR",
        "/upload",
        "str",
        "ftp",
        "FTP 根目录",
        description="远端父目录（会与 FTP_APP_DIR 拼接）；例如 /upload",
    )
)
_reg(
    SettingMeta(
        "FTP_APP_DIR",
        "aiprintword",
        "str",
        "ftp",
        "FTP 应用目录名",
        description="在 FTP_BASE_DIR 下的子目录名；默认 aiprintword",
    )
)
_reg(
    SettingMeta(
        "FTP_PASV",
        True,
        "bool",
        "ftp",
        "FTP 被动模式（PASV）",
        description="建议开启（移动/局域网/NAT 环境更稳定）；关闭则用主动模式",
    )
)


def ordered_keys() -> List[str]:
    return list(_ORDER)


def coerce_value(meta: SettingMeta, raw: str) -> Any:
    s = (raw or "").strip()
    t = meta.value_type
    if t == "bool":
        return str(s).strip().lower() not in ("0", "false", "no", "off", "")
    if t == "int":
        return int(float(s)) if s else int(meta.default)
    if t == "float":
        return float(s) if s else float(meta.default)
    return s if s else str(meta.default)


def meta_for(key: str) -> Optional[SettingMeta]:
    return REGISTRY.get(key)
