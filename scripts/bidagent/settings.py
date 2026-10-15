"""本地配置垫片：让从后端同步过来的服务模块在技能沙箱里也能拿到配置。

后端用 pydantic-settings 从 `.env` + 环境变量读配置；沙箱里既没有 `.env` 也不需要 pydantic，
所以这里用**纯标准库**实现同名属性：读**同名环境变量**、给**同样的默认值**，
默认值刻意对齐 `digital-hub-test/config/settings.py`，避免两处行为漂移。

沙箱里只有三个目录需要区别对待（其余都有环境变量可覆盖）：

    BID_OUTPUT_DIR   产物目录（Word / xlsx / png）       默认 /workspace/output
    BID_STATE_DIR    状态目录（检查会话/事实/台账/项目/模板）默认 <技能目录>/.state
    BID_INPUT_DIR    输入目录（资质清单与原件）            默认 /workspace/input

> 注意：状态目录默认放在**技能目录下**，是为了跨会话保留。若技能目录在会话里不可写
> （例如以只读方式挂载），会自动退到 `/workspace/.bid-state`（会话内有效）。
"""
from __future__ import annotations

import os
from pathlib import Path

# 技能目录：优先用 WeKnora 注入的变量，否则按本文件位置推断（<技能>/scripts/bidagent/settings.py）
SKILL_DIR = Path(os.environ.get("WEKNORA_SKILL_DIR") or Path(__file__).resolve().parents[2])


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


_STATE_DIR_CACHE: str | None = None


def _state_dir() -> str:
    """状态目录：优先技能目录，不可写则退回 /workspace。"""
    global _STATE_DIR_CACHE
    if _STATE_DIR_CACHE is None:
        for candidate in (SKILL_DIR / ".state", Path("/workspace/.bid-state")):
            try:
                candidate.mkdir(parents=True, exist_ok=True)
                _STATE_DIR_CACHE = str(candidate)
                break
            except OSError:
                continue
        else:
            _STATE_DIR_CACHE = str(SKILL_DIR / ".state")
    return _STATE_DIR_CACHE


class Settings:
    """与后端 `config/settings.py` 中 BID_* 字段同名同默认值的轻量替代。"""

    # ---------- 主体与文档 ----------
    BID_COMPANY_NAME = _env("BID_COMPANY_NAME", "广州优盟电子商务有限公司")
    BID_COMPANY_SHORT_NAME = _env("BID_COMPANY_SHORT_NAME", "优盟")

    # ---------- 目录 ----------
    BID_OUTPUT_DIR = _env("BID_OUTPUT_DIR", "/workspace/output")
    BID_UPLOAD_DIR = _env("BID_UPLOAD_DIR", "/workspace/input")
    BID_INPUT_DIR = _env("BID_INPUT_DIR", "/workspace/input")

    # ---------- 资质清单 ----------
    # 留空时由 bid_service.resolve_manifest_path 在输入目录里找 manifest.json
    BID_MANIFEST_PATH = _env("BID_MANIFEST_PATH", "")

    # ---------- 本地文档能力（沙箱里通常没有，留空即明确报错）----------
    BID_LIBREOFFICE = _env("BID_LIBREOFFICE", "")
    BID_OCR_COMMAND = _env("BID_OCR_COMMAND", "")
    BID_DOC_SERVICE_URL = _env("BID_DOC_SERVICE_URL", "")
    BID_DOC_SERVICE_TOKEN = _env("BID_DOC_SERVICE_TOKEN", "")
    BID_UPLOAD_MB = int(_env("BID_UPLOAD_MB", "80") or 80)
    BID_UPLOAD_EXT = _env("BID_UPLOAD_EXT", ".doc,.docx,.pdf,.txt,.md")

    # ---------- 内网 Dify（沙箱里不用；留空会回退本地启发式）----------
    BID_DIFY_URL = _env("BID_DIFY_URL", "")
    BID_ANALYSIS_KEY = _env("BID_ANALYSIS_KEY", "")
    BID_DRAFT_KEY = _env("BID_DRAFT_KEY", "")
    BID_DIFY_TIMEOUT = int(_env("BID_DIFY_TIMEOUT", "300") or 300)

    # ---------- 配图渲染 ----------
    BID_CHROMIUM_PATH = _env("BID_CHROMIUM_PATH", "")
    BID_FIGURE_BROWSER_CHANNEL = _env("BID_FIGURE_BROWSER_CHANNEL", "")
    BID_MERMAID_JS = _env("BID_MERMAID_JS", "")
    BID_FIGURE_TIMEOUT = int(_env("BID_FIGURE_TIMEOUT", "60") or 60)
    BID_FIGURE_WIDTH = int(_env("BID_FIGURE_WIDTH", "1240") or 1240)
    BID_FIGURE_SCALE = int(_env("BID_FIGURE_SCALE", "2") or 2)

    @property
    def bid_state_dir(self) -> str:
        """状态目录：环境变量优先，否则技能目录（不可写则退回 /workspace）。"""
        return _env("BID_STATE_DIR") or _state_dir()

    @property
    def bid_input_dir(self) -> str:
        return self.BID_INPUT_DIR or self.BID_UPLOAD_DIR

    @property
    def bid_upload_ext_list(self) -> list[str]:
        return [item.strip().lower() for item in self.BID_UPLOAD_EXT.split(",") if item.strip()]


settings = Settings()
