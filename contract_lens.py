#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ContractLens — 纯本地离线合同速读助手
========================================
专为每周审大量合同的产品经理设计。
上传合同 → 3秒输出决策卡片 → 一眼掌握全貌。

依赖安装:
    pip install streamlit pdfplumber python-docx reportlab pandas openpyxl

启动命令:
    streamlit run contract_lens.py

Author: ContractLens
Version: 1.0.0
"""

import io
import os
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

# ─── 第三方库 ───────────────────────────────────────────
import streamlit as st

# pandas 延迟加载（避免启动时 numpy C 扩展导致 segfault）
class _LazyPandas:
    _mod = None
    def __getattr__(self, name):
        if self._mod is None:
            import pandas as _p
            self._mod = _p
        return getattr(self._mod, name)

pd = _LazyPandas()

# ─── 页面配置（仅 streamlit run 时生效） ─────────────────
try:
    st.set_page_config(
        page_title="ContractLens · 合同速读",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
except Exception:
    pass  # 非 streamlit 环境下安全跳过


def _inject_css():
    """注入自定义 CSS 样式。"""
    try:
        st.markdown(
        """
<style>
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
    }
    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}

    .metric-card {
        background: #ffffff;
        border: 1px solid #e8ecf1;
        border-radius: 12px;
        padding: 20px 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        transition: box-shadow 0.2s;
    }
    .metric-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
    .metric-card .label {
        font-size: 12px; color: #8c939d; text-transform: uppercase;
        letter-spacing: 0.5px; margin-bottom: 6px;
    }
    .metric-card .value {
        font-size: 18px; font-weight: 700; color: #1a1f36; word-break: break-all;
    }
    .metric-card .value.amount { font-size: 22px; color: #e8543e; }

    .risk-high { background: #fef2f2; border-left: 4px solid #dc2626; padding: 12px 16px; border-radius: 0 8px 8px 0; margin: 8px 0; }
    .risk-medium { background: #fffbeb; border-left: 4px solid #f59e0b; padding: 12px 16px; border-radius: 0 8px 8px 0; margin: 8px 0; }
    .risk-low { background: #f0fdf4; border-left: 4px solid #22c55e; padding: 12px 16px; border-radius: 0 8px 8px 0; margin: 8px 0; }
    .risk-quote { font-size: 13px; color: #6b7280; margin-top: 6px; padding-left: 8px; border-left: 2px solid #d1d5db; }

    .summary-box {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px;
        padding: 20px 24px; line-height: 1.8; font-size: 14px; color: #334155;
    }
    .file-header {
        display: flex; align-items: center; gap: 12px;
        padding: 16px 0; border-bottom: 1px solid #f1f5f9; margin-bottom: 20px;
    }
    .file-header .file-icon { font-size: 32px; }
    .file-header .file-name { font-size: 18px; font-weight: 600; color: #1e293b; }
    .file-header .file-meta { font-size: 12px; color: #94a3b8; }
</style>
""",
        unsafe_allow_html=True,
    )
    except Exception:
        pass  # 非 streamlit 环境下安全跳过


# ═══════════════════════════════════════════════════════════
#  第〇部分：审查历史管理
# ═══════════════════════════════════════════════════════════

import json
import shutil

HISTORY_DIR = Path(__file__).parent / "history"


def _ensure_history_dir() -> Path:
    """确保历史记录目录存在。"""
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    return HISTORY_DIR


def save_review(
    filename: str,
    fields: dict[str, Any],
    risks: list[dict[str, Any]],
    summary: str,
    text_preview: str = "",
) -> str:
    """保存审查结果，返回记录 ID。"""
    _ensure_history_dir()
    record_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    record = {
        "id": record_id,
        "filename": filename,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "fields": fields,
        "risks": risks,
        "summary": summary,
        "text_preview": text_preview[:300],
        "risk_count": len(risks),
        "high_risk": sum(1 for r in risks if "高风险" in r.get("severity", "")),
        "mid_risk": sum(1 for r in risks if "中风险" in r.get("severity", "")),
        "low_risk": sum(1 for r in risks if "注意" in r.get("severity", "")),
    }
    filepath = _ensure_history_dir() / f"{record_id}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)
    return record_id


def load_all_reviews() -> list[dict[str, Any]]:
    """加载所有历史审查记录，按时间倒序。"""
    _ensure_history_dir()
    records: list[dict[str, Any]] = []
    for fp in sorted(HISTORY_DIR.glob("*.json"), reverse=True):
        try:
            with open(fp, "r", encoding="utf-8") as f:
                records.append(json.load(f))
        except (json.JSONDecodeError, OSError):
            pass
    return records


def load_review(record_id: str) -> dict[str, Any] | None:
    """加载单条审查记录。"""
    fp = _ensure_history_dir() / f"{record_id}.json"
    if fp.exists():
        with open(fp, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def delete_review(record_id: str) -> bool:
    """删除一条审查记录。"""
    fp = HISTORY_DIR / f"{record_id}.json"
    if fp.exists():
        fp.unlink()
        return True
    return False


def export_all_reviews() -> bytes:
    """导出全部历史记录为 JSON 文件。"""
    records = load_all_reviews()
    return json.dumps(records, ensure_ascii=False, indent=2).encode("utf-8")


def import_reviews(json_bytes: bytes) -> int:
    """从 JSON 导入历史记录，返回导入条数。"""
    _ensure_history_dir()
    try:
        data = json.loads(json_bytes.decode("utf-8"))
        if isinstance(data, list):
            count = 0
            for record in data:
                rid = record.get("id")
                if rid:
                    fp = HISTORY_DIR / f"{rid}.json"
                    if not fp.exists():
                        with open(fp, "w", encoding="utf-8") as f:
                            json.dump(record, f, ensure_ascii=False, indent=2)
                        count += 1
            return count
        elif isinstance(data, dict):
            rid = data.get("id")
            if rid:
                fp = HISTORY_DIR / f"{rid}.json"
                if not fp.exists():
                    with open(fp, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                    return 1
        return 0
    except (json.JSONDecodeError, UnicodeDecodeError):
        return 0


# ═══════════════════════════════════════════════════════════
#  第〇·五部分：邮件发送
# ═══════════════════════════════════════════════════════════

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# 常见邮箱 SMTP 配置
SMTP_PRESETS: dict[str, dict[str, Any]] = {
    "QQ邮箱": {"server": "smtp.qq.com", "port": 465, "use_ssl": True},
    "163邮箱": {"server": "smtp.163.com", "port": 465, "use_ssl": True},
    "126邮箱": {"server": "smtp.126.com", "port": 465, "use_ssl": True},
    "Gmail": {"server": "smtp.gmail.com", "port": 587, "use_ssl": False},
    "Outlook": {"server": "smtp.office365.com", "port": 587, "use_ssl": False},
    "自定义": {"server": "", "port": 465, "use_ssl": True},
}


def send_email_with_excel(
    smtp_server: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    recipient_email: str,
    subject: str,
    body: str,
    excel_data: bytes,
    filename: str,
    use_ssl: bool = True,
) -> tuple[bool, str]:
    """通过 SMTP 发送带 Excel 附件的邮件。返回 (成功, 消息)。"""
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # 附件
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{filename}"',
        )
        msg.attach(part)

        # 发送
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=15)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=15)
            server.starttls()

        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [recipient_email], msg.as_string())
        server.quit()
        return True, "邮件发送成功"
    except smtplib.SMTPAuthenticationError:
        return False, "认证失败：请检查邮箱地址和授权码是否正确"
    except smtplib.SMTPConnectError:
        return False, "连接失败：无法连接到 SMTP 服务器，请检查服务器地址和端口"
    except smtplib.SMTPException as e:
        return False, f"SMTP 错误：{e}"
    except Exception as e:
        return False, f"发送失败：{e}"


def generate_review_excel(fields: dict[str, Any], risks: list[dict[str, Any]], filename: str = "") -> bytes:
    """生成单份审查 Excel 报告。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_fields = pd.DataFrame([{"字段": k, "提取结果": v} for k, v in fields.items()])
        df_fields.to_excel(writer, sheet_name="合同字段", index=False)
        if risks:
            df_risks = pd.DataFrame([{
                "严重程度": r.get("severity", ""),
                "类别": r.get("category", ""),
                "说明": r.get("description", ""),
                "原文摘录": r.get("quote", ""),
            } for r in risks])
            df_risks.to_excel(writer, sheet_name="风险条款", index=False)
    buf.seek(0)
    return buf.read()


def generate_history_excel(records: list[dict[str, Any]]) -> bytes:
    """生成审查历史汇总 Excel。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # 汇总表
        summary_rows = []
        for r in records:
            fields = r.get("fields", {})
            summary_rows.append({
                "文件名": r.get("filename", ""),
                "审查时间": r.get("timestamp", ""),
                "甲方": fields.get("甲方", "")[:30],
                "乙方": fields.get("乙方", "")[:30],
                "合同金额": fields.get("合同金额", ""),
                "合同期限": fields.get("合同期限", "")[:30],
                "签署日期": fields.get("签署日期", ""),
                "生效条件": fields.get("生效条件", ""),
                "付款节奏": fields.get("付款节奏", "")[:50],
                "违约金": fields.get("违约金", "")[:50],
                "争议解决地": fields.get("争议解决地", ""),
                "保密期限": fields.get("保密期限", ""),
                "风险总数": r.get("risk_count", 0),
                "高风险": r.get("high_risk", 0),
                "中风险": r.get("mid_risk", 0),
                "低风险": r.get("low_risk", 0),
            })
        df_summary = pd.DataFrame(summary_rows)
        df_summary.to_excel(writer, sheet_name="审查汇总", index=False)

        # 每条详细记录单独一个 sheet（只做前20条避免文件过大）
        for i, r in enumerate(records[:20]):
            sheet_name = f"详情_{i+1}"[:31]  # Excel sheet name limit
            fields = r.get("fields", {})
            df_detail = pd.DataFrame([{"字段": k, "提取结果": v} for k, v in fields.items()])
            df_detail.to_excel(writer, sheet_name=sheet_name, index=False)

    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════
#  第〇·六部分：邮件设置 UI 组件
# ═══════════════════════════════════════════════════════════

def render_email_settings() -> dict[str, Any]:
    """渲染邮件设置表单，返回配置字典。"""
    with st.expander("📧 邮件发送设置（点击填写）", expanded=True):
        preset = st.selectbox("邮箱类型", list(SMTP_PRESETS.keys()), index=0)

        cfg = SMTP_PRESETS[preset]

        col1, col2 = st.columns(2)
        with col1:
            if preset == "自定义":
                server = st.text_input("SMTP 服务器", value=cfg["server"],
                                       placeholder="smtp.qq.com")
            else:
                server = cfg["server"]
                st.text_input("SMTP 服务器", value=server, disabled=True)
        with col2:
            port = st.number_input("端口", value=cfg["port"], min_value=1, max_value=65535)
        use_ssl = cfg["use_ssl"]

        sender = st.text_input("发件邮箱", placeholder="your_email@qq.com")
        password = st.text_input("授权码（非邮箱密码）", type="password",
                                 placeholder="QQ邮箱→设置→账户→POP3/SMTP服务→生成授权码")
        recipient = st.text_input("收件邮箱", placeholder="recipient@example.com")

        return {
            "server": server,
            "port": port,
            "use_ssl": use_ssl,
            "sender": sender,
            "password": password,
            "recipient": recipient,
        }


def validate_email_settings(cfg: dict[str, Any]) -> list[str]:
    """验证邮件配置，返回缺失项列表。"""
    missing = []
    if not cfg.get("server"):
        missing.append("SMTP 服务器")
    if not cfg.get("sender"):
        missing.append("发件邮箱")
    if not cfg.get("password"):
        missing.append("授权码")
    if not cfg.get("recipient"):
        missing.append("收件邮箱")
    return missing


# ═══════════════════════════════════════════════════════════
#  第一部分：文本提取
# ═══════════════════════════════════════════════════════════

def extract_text_from_pdf(file_bytes: bytes) -> str:
    """从 PDF 中提取纯文本（pdfplumber）。"""
    import pdfplumber

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages_text: list[str] = []
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages_text.append(text)
            result = "\n".join(pages_text)
            if not result.strip():
                return "[提示] 此 PDF 可能为扫描件，暂无文字层，请使用 OCR 处理后再试。"
            return result
    except Exception as e:
        return f"[错误] PDF 解析失败: {e}"


def extract_text_from_docx(file_bytes: bytes) -> str:
    """从 DOCX 中提取纯文本。"""
    import docx

    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        return f"[错误] DOCX 解析失败: {e}"


def extract_text_from_txt(file_bytes: bytes) -> str:
    """从 TXT 中提取文本，尝试多种编码。"""
    for enc in ["utf-8", "gbk", "gb2312", "gb18030", "latin-1"]:
        try:
            return file_bytes.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return file_bytes.decode("utf-8", errors="replace")


def extract_text(file_bytes: bytes, filename: str) -> str:
    """根据文件类型自动选择提取方式。"""
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_bytes)
    elif ext in (".docx", ".doc"):
        return extract_text_from_docx(file_bytes)
    elif ext == ".txt":
        return extract_text_from_txt(file_bytes)
    else:
        return f"[错误] 不支持的文件格式: {ext}"


# ═══════════════════════════════════════════════════════════
#  第二部分：中文数字 → 阿拉伯数字
# ═══════════════════════════════════════════════════════════

CN_NUM_MAP: dict[str, int] = {
    "零": 0, "〇": 0,
    "一": 1, "壹": 1,
    "二": 2, "两": 2, "贰": 2,
    "三": 3, "叁": 3,
    "四": 4, "肆": 4,
    "五": 5, "伍": 5,
    "六": 6, "陆": 6,
    "七": 7, "柒": 7,
    "八": 8, "捌": 8,
    "九": 9, "玖": 9,
    "十": 10, "拾": 10,
    "百": 100, "佰": 100,
    "千": 1000, "仟": 1000,
    "万": 10_000, "萬": 10_000,
    "亿": 100_000_000, "億": 100_000_000,
}

CN_DIGITS = {"零", "〇", "一", "二", "两", "三", "四", "五", "六", "七", "八", "九",
             "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"}
CN_POWERS = {"十": 10, "拾": 10, "百": 100, "佰": 100, "千": 1000, "仟": 1000}


def _parse_cn_segment(seg: str) -> int:
    """解析不含「万/亿」的纯中文数字片段，如 '一百二十三' → 123。"""
    total = 0
    current = 0
    for ch in seg:
        if ch in CN_DIGITS:
            current = CN_NUM_MAP[ch]
        elif ch in CN_POWERS:
            unit = CN_POWERS[ch]
            if current == 0:
                current = 1  # "十" 等同于 "一十"
            total += current * unit
            current = 0
        else:
            # 非数字字符跳过
            pass
    total += current
    return total


def cn_numeral_to_int(text: str) -> int | None:
    """
    将中文大写/小写金额字符串转为整数。
    例: "壹佰贰拾叁万肆仟伍佰陆拾柒" → 1234567
        "一百二十三万四千五百六十七" → 1234567
    """
    # 清理前缀
    text = re.sub(r"人民币|美元|港币|欧元|日元", "", text)
    text = re.sub(r"元整|元|整", "", text)
    text = text.strip()
    if not text:
        return None

    # 检查是否全是中文数字字符
    for ch in text:
        if ch not in CN_NUM_MAP:
            return None  # 不是纯中文数字

    # 分段：按「亿」「万」拆分
    yi_part = re.split(r"[亿億]", text, maxsplit=1)
    if len(yi_part) == 2:
        yi_val = _parse_cn_segment(yi_part[0])
        rest = yi_part[1]
    else:
        yi_val = 0
        rest = yi_part[0]

    wan_part = re.split(r"[万萬]", rest, maxsplit=1)
    if len(wan_part) == 2:
        wan_val = _parse_cn_segment(wan_part[0])
        ge_val = _parse_cn_segment(wan_part[1])
    else:
        wan_val = 0
        ge_val = _parse_cn_segment(wan_part[0])

    return yi_val * 100_000_000 + wan_val * 10_000 + ge_val


def extract_amount(text: str) -> str | None:
    """
    从文本中提取合同金额，支持：
    - 阿拉伯数字 + 单位: "100万元", "¥1,234,567.89"
    - 中文大写:  "人民币壹佰万元整"
    - 返回归一化数字字符串
    """

    # 策略1：匹配"数字 + 万元/元"模式
    patterns = [
        # 100万元 / 100万元整
        r"([\d,]+\.?\d*)\s*万\s*元",
        # ¥1,234,567.89 / 人民币1,234,567.89元
        r"(?:人民币|¥|￥|RMB|CNY)?\s*([\d,]+\.?\d{0,2})\s*元",
        # $1,000,000
        r"\$\s*([\d,]+\.?\d{0,2})",
    ]

    for pat in patterns:
        m = re.search(pat, text)
        if m:
            num_str = m.group(1).replace(",", "")
            try:
                val = float(num_str)
                if "万" in m.group(0):
                    val *= 10_000
                # 格式化为千分位
                if val >= 10_000:
                    return f"¥{val:,.0f}（{val/10000:g}万元）"
                else:
                    return f"¥{val:,.2f}"
            except ValueError:
                pass

    # 策略2：中文大写金额
    cn_pattern = r"((?:人民币|美元|港币)?[壹贰叁肆伍陆柒捌玖拾佰仟万亿兩零一二三四五六七八九十百千万亿]{2,40}(?:元整|元|整)?)"
    for m in re.finditer(cn_pattern, text):
        result = cn_numeral_to_int(m.group(1))
        if result is not None and result > 0:
            if result >= 10_000:
                return f"¥{result:,}（{result/10000:g}万元）"
            else:
                return f"¥{result:,}"

    return None


# ═══════════════════════════════════════════════════════════
#  第三部分：关键字段提取
# ═══════════════════════════════════════════════════════════

def safe_search(pattern: str, text: str, group: int = 0) -> str | None:
    """执行正则搜索，异常时返回 None。"""
    try:
        m = re.search(pattern, text, re.DOTALL)
        if m:
            try:
                return m.group(group).strip()
            except IndexError:
                return m.group(0).strip()  # 指定组不存在时回退到完整匹配
    except re.error:
        pass
    return None


def extract_parties(text: str) -> dict[str, str]:
    """提取甲方、乙方信息。"""
    parties: dict[str, str] = {}

    # 甲方
    jia_patterns = [
        r"甲\s*方[（(]?[^)）：:：\n]{0,20}[)）]?\s*[：:：]\s*([^\n]{1,80})",
        r"购买方[：:：]\s*([^\n]{1,80})",
        r"委托方[：:：]\s*([^\n]{1,80})",
        r"发包方[：:：]\s*([^\n]{1,80})",
        r"定作方[：:：]\s*([^\n]{1,80})",
    ]
    for pat in jia_patterns:
        val = safe_search(pat, text, 1)
        if val:
            parties["甲方"] = val
            break
    if "甲方" not in parties:
        parties["甲方"] = "未提取"

    # 乙方
    yi_patterns = [
        r"乙\s*方[（(]?[^)）：:：\n]{0,20}[)）]?\s*[：:：]\s*([^\n]{1,80})",
        r"供应方[：:：]\s*([^\n]{1,80})",
        r"受托方[：:：]\s*([^\n]{1,80})",
        r"承包方[：:：]\s*([^\n]{1,80})",
        r"承揽方[：:：]\s*([^\n]{1,80})",
    ]
    for pat in yi_patterns:
        val = safe_search(pat, text, 1)
        if val:
            parties["乙方"] = val
            break
    if "乙方" not in parties:
        # 如果只有一个主体，尝试找"供应商"等
        alt = safe_search(r"(?:供应商|服务方|开发方|设计方)[：:：]\s*([^\n]{1,80})", text, 1)
        parties["乙方"] = alt if alt else "未提取"

    return parties


def extract_sign_date(text: str) -> str:
    """提取签署日期。"""
    patterns = [
        r"签署日期[：:：]\s*(\d{4}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}\s*日?)",
        r"签订日期[：:：]\s*(\d{4}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}\s*日?)",
        r"签约日期[：:：]\s*(\d{4}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}\s*日?)",
        r"日期[：:：]\s*(\d{4}\s*[-/年]\s*\d{1,2}\s*[-/月]\s*\d{1,2}\s*日?)",
        r"(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)",
        r"(\d{4}[-/]\d{1,2}[-/]\d{1,2})",
    ]
    for pat in patterns:
        val = safe_search(pat, text, 1)
        if val:
            return val
    return "未提取"


def extract_effective_condition(text: str) -> str:
    """提取生效条件。"""
    patterns = [
        r"(签字.*?盖章.*?之日起?生效)",
        r"(双方.*?盖章.*?后?生效)",
        r"(自.*?之日起?生效[。；;]?)",
        r"生效[：:：]?\s*([^\n。；;]{2,60})",
    ]
    for pat in patterns:
        val = safe_search(pat, text, 1)
        if val and len(val) > 4:
            return val
    return "未提取"


def extract_contract_term(text: str) -> dict[str, str]:
    """提取合同期限与自动续约。"""
    result: dict[str, str] = {"期限": "未提取", "自动续约": "未检测到"}

    # 期限起止
    term_patterns = [
        r"合同期限[：:：]?\s*([^\n。；;]{2,80})",
        r"履行期限[：:：]?\s*([^\n。；;]{2,80})",
        r"合作期限[：:：]?\s*([^\n。；;]{2,80})",
        r"(自\s*\d{4}.*?(?:止|至|到)\s*\d{4}.*?[日。；;])",
        r"期限[：:：]?\s*(\d+\s*[个]?\s*[年月日周])",
        r"(合同.*?有效.*?[^\n。；;]{2,60})",
    ]
    for pat in term_patterns:
        val = safe_search(pat, text, 1)
        if val and len(val) > 2:
            result["期限"] = val
            break

    # 自动续约
    renewal_pat = r"(自动续[约期签]|自动延[期长]|自动顺延|视为续[约签]|到期.*?未.*?提出.*?视为.*?续)"
    if re.search(renewal_pat, text):
        # 提取完整句子
        sent = safe_search(r"([^\n。；;]*自动续[约期签][^\n。；;]*[。；;]?)", text, 1)
        result["自动续约"] = sent if sent else "是（存在自动续约条款）"
    else:
        result["自动续约"] = "未检测到"

    return result


def extract_payment(text: str) -> str:
    """提取付款节奏。"""
    # 先找"付款"段落
    pay_section = safe_search(r"(?:付款[^\n]*|支付[^\n]*)[：:：]?\s*([^\n]{5,200})", text, 1)
    if pay_section:
        return pay_section[:200]

    # 找分期描述
    phases_pat = r"分\s*[一二三四五六七八九十\d]+\s*期"
    if re.search(phases_pat, text):
        ctx = safe_search(r"([^\n。；;]*分\s*[一二三四五六七八九十\d]+\s*期[^\n。；;]*[。；;]?)", text, 1)
        if ctx:
            return ctx[:200]

    # 找比例描述
    pct_pat = r"(\d{1,3})\s*%\s*(?:.*?付款|.*?支付)"
    matches = re.findall(pct_pat, text)
    if matches:
        return f"检测到 {len(matches)} 个比例付款节点：{', '.join(m + '%' for m in matches)}"

    return "未提取"


def extract_deliverables(text: str) -> str:
    """提取交付物/服务范围（一句话概括）。"""
    patterns = [
        r"(?:交付物|交付内容|服务内容|工作内容|服务范围|项目范围|合同标的)[：:：]\s*([^\n。；;]{4,200})",
        r"(?:乙方|受托方|承包方)\s*(?:负责|完成|提供|交付|开发|设计)\s*([^\n。；;]{6,120})",
        r"标\s*的[：:：]\s*([^\n。；;]{6,200})",
    ]
    for pat in patterns:
        val = safe_search(pat, text, 1)
        if val and len(val) > 4 and val not in ("第一条 合同标的",):
            return val[:200]
    return "未提取"


def extract_penalty(text: str) -> str:
    """提取违约金信息。"""
    patterns = [
        # 违约金明确说明（带冒号）
        r"违约金[：:：]\s*([^\n。；;]{4,200})",
        # X%的违约金 或 违约金...X%（百分比可前可后）
        r"([^\n。；;]{0,80}(?:[\d.]+%.*?违约金|违约金.*?[\d.]+%)[^\n。；;]{0,50})",
        # 按日/按天计算违约金
        r"(按\s*(?:每日|每天).*?[0-9.]+%.*?违约金)",
        # 违约责任段落
        r"违约责任[：:：]?\s*([^\n]{5,300})",
    ]
    for pat in patterns:
        val = safe_search(pat, text, 1)
        if val and len(val) > 3:
            return val[:250]

    return "未提取"


def extract_dispute_resolution(text: str) -> str:
    """提取争议解决地。"""
    patterns = [
        r"(?:争议|仲裁|管辖).*?(?:提交|由|向)\s*([^\n。；;]{3,80}?(?:法院|仲裁委|仲裁委员会|仲裁机构))",
        r"(?:管辖.*?法院|所在地.*?法院)\s*[：:：]?\s*([^\n。；;]{2,80})",
        r"仲裁.*?([^\n。；;]{4,80}?(?:仲裁委|仲裁委员会))",
        r"争议.*?解决[：:：]?\s*([^\n。；;]{4,200})",
    ]
    for pat in patterns:
        val = safe_search(pat, text, 1)
        if val:
            return val[:200]
    return "未提取"


def extract_confidentiality(text: str) -> str:
    """提取保密期限。"""
    # 永久保密
    if re.search(r"(永久.*?保密|保密.*?永久|无期限.*?保密|保密.*?无期限)", text):
        return "永久"

    # 合同终止后X年
    pat = r"(?:合同.*?终止|合同.*?解除|合同.*?期满).*?[后起]\s*(\d+)\s*[年]"
    m = re.search(pat, text)
    if m:
        return f"合同终止后{m.group(1)}年"

    # 保密期限
    conf_pat = r"保密.*?(?:期限|期[间限])[：:：]?\s*([^\n。；;]{3,80})"
    val = safe_search(conf_pat, text, 1)
    if val:
        return val[:200]

    return "未提取"


def extract_all_fields(text: str) -> dict[str, Any]:
    """一次性提取所有关键字段。"""
    parties = extract_parties(text)
    amount = extract_amount(text)
    sign_date = extract_sign_date(text)
    effective = extract_effective_condition(text)
    term = extract_contract_term(text)
    payment = extract_payment(text)
    deliverables = extract_deliverables(text)
    penalty = extract_penalty(text)
    dispute = extract_dispute_resolution(text)
    confidentiality = extract_confidentiality(text)

    return {
        "甲方": parties.get("甲方", "未提取"),
        "乙方": parties.get("乙方", "未提取"),
        "合同金额": amount or "未提取",
        "签署日期": sign_date,
        "生效条件": effective,
        "合同期限": term["期限"],
        "自动续约": term["自动续约"],
        "付款节奏": payment,
        "交付物/服务范围": deliverables,
        "违约金": penalty,
        "争议解决地": dispute,
        "保密期限": confidentiality,
    }


# ═══════════════════════════════════════════════════════════
#  第四部分：风险扫描
# ═══════════════════════════════════════════════════════════

RISK_RULES: list[dict[str, Any]] = [
    # 🔴 高风险
    {
        "category": "单方免责/解约权",
        "keywords": [
            r"单方解除", r"单方终止", r"单方面.*?解除", r"单方.*?变更",
            r"不承担任何责任", r"概不负责", r"免除.*?一切.*?责任",
            r"不予赔偿", r"不承担.*?赔偿",
        ],
        "severity": "🔴 高风险",
        "css_class": "risk-high",
        "description": "赋予单方过大的免责或解约权利，可能导致对方任意终止合作",
    },
    {
        "category": "无限/加重责任",
        "keywords": [
            r"无限责任", r"无条件.*?承担", r"连带责任",
            r"承担.*?一切.*?损失", r"全部.*?赔偿",
            r"惩罚性.*?赔偿", r"间接损失",
        ],
        "severity": "🔴 高风险",
        "css_class": "risk-high",
        "description": "不合理的责任加重条款，可能导致承担超出预期的赔偿责任",
    },
    {
        "category": "知识产权归属风险",
        "keywords": [
            r"知识产权.*?归.*?对方",
            r"知识产权.*?无偿.*?转让",
            r"放弃.*?知识产权",
            r"知识产权.*?归属.*?另行.*?约定",
            r"知识产权.*?未.*?约定",
            r"知识产权.*?归属.*?不[明确清]",
        ],
        "severity": "🟠 中风险",
        "css_class": "risk-medium",
        "description": "知识产权归属存在不确定性或对己方可能不利",
    },
    # 🟠 中风险
    {
        "category": "模糊/不确定条款",
        "keywords": [
            r"合理.*?(?:期限|时间|费用|价格|数量)",
            r"及时.*?(?:通知|支付|交付|处理)",
            r"适[当量]", r"酌[情定]", r"视情况",
            r"必要.*?时", r"尽[快力]", r"争取",
        ],
        "severity": "🟠 中风险",
        "css_class": "risk-medium",
        "description": "使用模糊词汇，缺乏明确的量化标准，容易产生争议",
    },
    {
        "category": "管辖地不利",
        "keywords": [
            r"对方.*?所在地.*?(?:法院|仲裁)",
            r"被告.*?所在地.*?(?:法院|管辖)",
            r"仲裁.*?地.*?(?:对方|北京|上海|广州|深圳)",
        ],
        "severity": "🟠 中风险",
        "css_class": "risk-medium",
        "description": "争议解决地对己方不利，增加维权成本",
    },
    {
        "category": "违约金过高",
        "keywords": [
            r"违约金.*?[3-9]\d%",
            r"每日.*?[0-9.]+%.*?违约金",
            r"按日.*?累计.*?违约",
            r"违约金.*?\d{2,}%",
        ],
        "severity": "🟠 中风险",
        "css_class": "risk-medium",
        "description": "违约金比例可能过高，超出合理范围",
    },
    {
        "category": "付款条件苛刻",
        "keywords": [
            r"验收.*?合格.*?后.*?\d+.*?[天日].*?付款",
            r"全部.*?交付.*?后.*?付款",
            r"终[验收].*?后.*?支付",
            r"收到.*?款项.*?后.*?支付",
            r"先.*?开票.*?后.*?付款",
        ],
        "severity": "🟠 中风险",
        "css_class": "risk-medium",
        "description": "付款条件对收款方不利，可能造成现金流压力",
    },
    # 🟡 注意
    {
        "category": "自动续约陷阱",
        "keywords": [
            r"自动续[约期签]", r"自动延[期长]", r"自动顺延",
            r"视为续[签约]", r"到期.*?未.*?提出.*?视为.*?续",
            r"合同.*?期满.*?自动",
        ],
        "severity": "🟡 注意",
        "css_class": "risk-low",
        "description": "存在自动续约条款，容易被忽略，导致被动续约",
    },
    {
        "category": "保密期限过长",
        "keywords": [
            r"永久.*?保密", r"保密.*?永久",
            r"保密.*?无限[期制]", r"保密.*?无.*?期限",
            r"保密.*?终止.*?后.*?(?:10|[2-9]\d?)\s*年",
        ],
        "severity": "🟡 注意",
        "css_class": "risk-low",
        "description": "保密义务期限过长或无限期，可能不合理地限制未来发展",
    },
    {
        "category": "不可抗力范围过宽",
        "keywords": [
            r"不可抗力.*?包括", r"不可抗力.*?不限于",
            r"政策.*?变化.*?不可抗力",
            r"市场.*?变化.*?不可抗力",
        ],
        "severity": "🟡 注意",
        "css_class": "risk-low",
        "description": "不可抗力定义范围过宽，可能被滥用",
    },
]


def scan_risks(text: str) -> list[dict[str, Any]]:
    """扫描并返回风险列表。"""
    findings: list[dict[str, Any]] = []

    for rule in RISK_RULES:
        for kw_pat in rule["keywords"]:
            try:
                matches = list(re.finditer(kw_pat, text))
                for m in matches:
                    # 提取包含匹配词的完整句子作为引用
                    match_pos = m.start()
                    match_end = m.end()
                    # 找到匹配词前的句子边界
                    before = text[max(0, match_pos - 120):match_pos]
                    last_sep = 0
                    for sep in "。\n；;":
                        pos = before.rfind(sep)
                        if pos > last_sep:
                            last_sep = pos
                    sent_start = max(0, match_pos - 120) + last_sep + 1
                    # 找到匹配词后的句子边界
                    after = text[match_end:min(len(text), match_end + 120)]
                    first_sep = len(after)
                    for sep in "。\n；;":
                        pos = after.find(sep)
                        if 0 < pos < first_sep:
                            first_sep = pos
                    sent_end = match_end + first_sep + 1
                    snippet = text[sent_start:sent_end].replace("\n", " ").strip()
                    # 限制长度
                    if len(snippet) > 150:
                        snippet = snippet[:147] + "..."

                    findings.append({
                        "category": rule["category"],
                        "severity": rule["severity"],
                        "css_class": rule["css_class"],
                        "description": rule["description"],
                        "matched": m.group(0),
                        "quote": snippet,
                    })
            except re.error:
                continue

    # 去重：按匹配字符串去重
    seen: set[str] = set()
    unique_findings: list[dict[str, Any]] = []
    for f in findings:
        key = f"{f['category']}|{f['matched']}"
        if key not in seen:
            seen.add(key)
            unique_findings.append(f)

    return unique_findings


# ═══════════════════════════════════════════════════════════
#  第五部分：规则摘要生成
# ═══════════════════════════════════════════════════════════

def split_sentences(text: str) -> list[str]:
    """将文本切分为句子列表。"""
    # 按句号、分号、换行等分割
    raw = re.split(r"[。；;！!\n]+", text)
    return [s.strip() for s in raw if len(s.strip()) > 6]


def generate_summary(text: str, max_chars: int = 200) -> str:
    """
    规则提取生成合同摘要：
    1. 每段第一句（主旨句）
    2. 含「约定/承诺/保证/负责」的义务性句子
    3. 含「如/若/一旦/如果」的条件触发句
    """
    sentences = split_sentences(text)
    if not sentences:
        return "（未能提取到有效句子，请检查合同文本格式）"

    first_sentences: list[str] = []
    obligation_sentences: list[str] = []
    conditional_sentences: list[str] = []

    # 模拟段落：按双换行或条款编号分割
    paragraphs = re.split(r"\n\s*\n|第[一二三四五六七八九十\d]+条", text)
    for para in paragraphs:
        sents = split_sentences(para)
        if sents:
            first_sentences.append(sents[0])

    # 义务性句子
    obl_pat = re.compile(r"(约定|承诺|保证|负责|应当|必须|应按|须按|不得|禁止)")
    for s in sentences:
        if obl_pat.search(s) and s not in first_sentences:
            obligation_sentences.append(s)

    # 条件触发句
    cond_pat = re.compile(r"(^|[。；;！!\n])[^。；;！!\n]*(?:如|若|一旦|如果|除非|假如)[^。；;！!\n]*[。；;！!]?")
    for m in cond_pat.finditer(text):
        s = m.group(0).strip()
        if len(s) > 10 and s not in first_sentences and s not in obligation_sentences:
            conditional_sentences.append(s)

    # 拼合摘要
    summary_parts: list[str] = []
    char_count = 0

    # 优先取主旨句
    for s in first_sentences[:3]:
        if char_count + len(s) <= max_chars:
            summary_parts.append(s + "。")
            char_count += len(s) + 1

    # 取义务句
    for s in obligation_sentences[:3]:
        if char_count + len(s) <= max_chars:
            summary_parts.append(s + "。")
            char_count += len(s) + 1

    # 取条件句
    for s in conditional_sentences[:2]:
        if char_count + len(s) <= max_chars:
            summary_parts.append(s if s.endswith("。") else s + "。")
            char_count += len(s) + 1

    if not summary_parts:
        # 兜底：取前几句
        for s in sentences[:4]:
            if char_count + len(s) <= max_chars:
                summary_parts.append(s + "。")
                char_count += len(s) + 1

    return "".join(summary_parts) if summary_parts else "（无法生成摘要）"


# ═══════════════════════════════════════════════════════════
#  第六部分：PDF 导出
# ═══════════════════════════════════════════════════════════

def generate_pdf(
    filename: str,
    fields: dict[str, Any],
    risks: list[dict[str, Any]],
    summary: str,
) -> bytes | None:
    """使用 reportlab 生成 PDF 报告。reportlab 未安装时返回 None。"""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.colors import HexColor
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
        )
        from reportlab.pdfbase import pdfmetrics
    except ImportError:
        return None  # reportlab 未安装，跳过 PDF 生成

    # ── 跨平台中文字体检测 ──
    cn_font = "Helvetica"  # 兜底

    def _try_cid_font(name: str) -> str | None:
        """尝试注册 CID 字体（Windows/macOS 内置）。"""
        try:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont(name))
            return name
        except Exception:
            return None

    def _try_ttf_font(path: str, name: str) -> str | None:
        """尝试注册 TTF 字体文件。"""
        try:
            from reportlab.pdfbase.ttfonts import TTFont
            if Path(path).exists():
                pdfmetrics.registerFont(TTFont(name, path))
                return name
        except Exception:
            pass
        return None

    # 策略1: CID 字体（Windows/macOS reportlab 内置）
    for cid_name in ["STSong-Light", "HeiseiMin-W3", "HYGoThic-Medium"]:
        result = _try_cid_font(cid_name)
        if result:
            cn_font = result
            break

    # 策略2: Linux 常见 CJK 字体路径
    if cn_font == "Helvetica":
        linux_font_paths = [
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
            "/usr/share/fonts/opentype/noto/NotoSansSC-Regular.otf",
            "/usr/share/fonts/truetype/arphic/uming.ttc",
            "/usr/share/fonts/truetype/arphic/ukai.ttc",
        ]
        for font_path in linux_font_paths:
            result = _try_ttf_font(font_path, "CJKFont")
            if result:
                cn_font = result
                break

    # 策略3: macOS 系统字体路径
    if cn_font == "Helvetica":
        mac_font_paths = [
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/Library/Fonts/Arial Unicode.ttf",
        ]
        for font_path in mac_font_paths:
            result = _try_ttf_font(font_path, "CJKFont")
            if result:
                cn_font = result
                break

    # 策略4: Windows 系统字体
    if cn_font == "Helvetica":
        win_font_paths = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc",
            "C:/Windows/Fonts/simhei.ttf",
        ]
        for font_path in win_font_paths:
            result = _try_ttf_font(font_path, "CJKFont")
            if result:
                cn_font = result
                break

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=20 * mm, rightMargin=20 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CNTitle", parent=styles["Title"],
        fontName=cn_font, fontSize=16, leading=22,
        spaceAfter=6,
    )
    h2_style = ParagraphStyle(
        "CNH2", parent=styles["Heading2"],
        fontName=cn_font, fontSize=13, leading=18,
        spaceBefore=12, spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "CNBody", parent=styles["Normal"],
        fontName=cn_font, fontSize=9, leading=14,
    )
    small_style = ParagraphStyle(
        "CNSmall", parent=styles["Normal"],
        fontName=cn_font, fontSize=8, leading=12,
        textColor=HexColor("#6b7280"),
    )

    story: list[Any] = []

    # 标题
    story.append(Paragraph(f"ContractLens · 合同审查报告", title_style))
    story.append(Paragraph(f"文件: {filename} | 审查时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}", small_style))
    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=HexColor("#e2e8f0")))
    story.append(Spacer(1, 6 * mm))

    # 核心指标
    story.append(Paragraph("核心指标", h2_style))
    core_data = [
        ["甲方", fields.get("甲方", "—")],
        ["乙方", fields.get("乙方", "—")],
        ["合同金额", fields.get("合同金额", "—")],
        ["合同期限", fields.get("合同期限", "—")],
        ["签署日期", fields.get("签署日期", "—")],
    ]
    core_table = Table(core_data, colWidths=[80, 380])
    core_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), cn_font),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#f8fafc")),
        ("TEXTCOLOR", (0, 0), (0, -1), HexColor("#64748b")),
        ("GRID", (0, 0), (-1, -1), 0.3, HexColor("#e2e8f0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(core_table)
    story.append(Spacer(1, 6 * mm))

    # 全部字段
    story.append(Paragraph("详细字段", h2_style))
    detail_data = [[k, v] for k, v in fields.items()]
    detail_table = Table(detail_data, colWidths=[100, 360])
    detail_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), cn_font),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (0, -1), HexColor("#f8fafc")),
        ("GRID", (0, 0), (-1, -1), 0.3, HexColor("#e2e8f0")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(detail_table)
    story.append(Spacer(1, 6 * mm))

    # 风险条款
    if risks:
        story.append(Paragraph(f"风险扫描（共 {len(risks)} 条）", h2_style))
        for r in risks:
            risk_text = f"<b>{r['severity']} {r['category']}</b><br/>{r['quote']}"
            story.append(Paragraph(risk_text, body_style))
            story.append(Spacer(1, 2 * mm))
    else:
        story.append(Paragraph("风险扫描：未检测到明显风险条款", body_style))

    story.append(Spacer(1, 6 * mm))

    # 摘要
    story.append(Paragraph("合同摘要", h2_style))
    story.append(Paragraph(summary, body_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════
#  第七部分：单个合同 UI
# ═══════════════════════════════════════════════════════════

def render_single_file_ui():
    """单个合同审查界面。"""
    st.header("📄 单份合同审查")

    uploaded = st.file_uploader(
        "上传合同文件",
        type=["pdf", "docx", "doc", "txt"],
        key="single_upload",
        help="支持 PDF、DOCX、TXT 格式",
    )

    if not uploaded:
        # 空状态
        st.info("👆 上传一份合同文件，开始快速审查")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("#### 📎 支持格式")
            st.markdown("- PDF（文字层）\n- Word (.docx)\n- 纯文本 (.txt)")
        with col2:
            st.markdown("#### ⚡ 3秒出结果")
            st.markdown("- 9大关键字段提取\n- 风险条款自动扫描\n- 智能摘要生成")
        with col3:
            st.markdown("#### 🔒 完全离线")
            st.markdown("- 数据不出本地\n- 无需API密钥\n- 隐私安全无忧")
        return

    # 读取文件
    file_bytes = uploaded.read()
    text = extract_text(file_bytes, uploaded.name)

    if text.startswith("[错误]") or text.startswith("[提示]"):
        st.error(text)
        return

    # ── 处理 ──
    with st.spinner("🔍 正在分析合同..."):
        fields = extract_all_fields(text)
        risks = scan_risks(text)
        summary = generate_summary(text)

    # ── 决策卡片 ──
    st.markdown("---")

    # 文件头部
    st.markdown(f"""
    <div class="file-header">
        <div class="file-icon">📄</div>
        <div>
            <div class="file-name">{uploaded.name}</div>
            <div class="file-meta">审查时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} · 文本长度: {len(text):,} 字符</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 四大核心指标卡片 ──
    col_a, col_b, col_c, col_d = st.columns(4)

    with col_a:
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">甲方</div>
            <div class="value" style="font-size:15px;">{fields['甲方'][:30]}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_b:
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">乙方</div>
            <div class="value" style="font-size:15px;">{fields['乙方'][:30]}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_c:
        amount_display = fields['合同金额'].replace("¥", "") if fields['合同金额'] != "未提取" else "未提取"
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">合同金额</div>
            <div class="value amount">{amount_display}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_d:
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">合同期限</div>
            <div class="value" style="font-size:14px;">{fields['合同期限'][:30]}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 详细字段表格 ──
    with st.expander("📋 全部提取字段", expanded=True):
        detail_rows = []
        for k, v in fields.items():
            # 跳过已在卡片中展示的核心字段
            detail_rows.append({"字段": k, "提取结果": v})
        st.dataframe(
            pd.DataFrame(detail_rows),
            use_container_width=True,
            hide_index=True,
            column_config={
                "字段": st.column_config.TextColumn("字段", width="small"),
                "提取结果": st.column_config.TextColumn("提取结果", width="large"),
            },
        )

    # ── 风险扫描 ──
    risk_count = len(risks)
    high_count = sum(1 for r in risks if "高风险" in r["severity"])
    mid_count = sum(1 for r in risks if "中风险" in r["severity"])
    low_count = sum(1 for r in risks if "注意" in r["severity"])

    with st.expander(f"⚠️ 风险扫描（共 {risk_count} 条 · 🔴{high_count} 🟠{mid_count} 🟡{low_count}）", expanded=True):
        if not risks:
            st.success("✅ 未检测到明显风险条款")
        else:
            for i, risk in enumerate(risks):
                st.markdown(f"""
                <div class="{risk['css_class']}">
                    <strong>{risk['severity']} · {risk['category']}</strong>
                    <div style="font-size:12px;color:#6b7280;margin-top:2px;">{risk['description']}</div>
                    <div class="risk-quote">📌 原文: 「{risk['quote']}」</div>
                </div>
                """, unsafe_allow_html=True)

    # ── 摘要 ──
    with st.expander("📝 合同摘要", expanded=True):
        st.markdown(f'<div class="summary-box">{summary}</div>', unsafe_allow_html=True)

    # ── 导出按钮 ──
    st.markdown("---")
    col_exp1, col_exp2, col_exp3, col_exp4 = st.columns([1, 1, 1, 3])
    with col_exp1:
        pdf_data = generate_pdf(uploaded.name, fields, risks, summary)
        if pdf_data is not None:
            st.download_button(
                label="📥 导出 PDF",
                data=pdf_data,
                file_name=f"ContractLens_{Path(uploaded.name).stem}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
    with col_exp2:
        # Excel 导出
        df_export = pd.DataFrame([
            {"字段": k, "提取结果": v} for k, v in fields.items()
        ])
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df_export.to_excel(writer, sheet_name="合同字段", index=False)
            if risks:
                pd.DataFrame(risks).to_excel(writer, sheet_name="风险条款", index=False)
        excel_buf.seek(0)
        st.download_button(
            label="📥 导出 Excel",
            data=excel_buf,
            file_name=f"ContractLens_{Path(uploaded.name).stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── 发送邮件 ──
    st.markdown("---")
    email_cfg = render_email_settings()
    col_em1, col_em2, _ = st.columns([1, 1, 4])
    with col_em1:
        if st.button("📧 发送 Excel 到邮箱", use_container_width=True,
                     type="primary"):
            missing = validate_email_settings(email_cfg)
            if missing:
                st.error(f"请填写：{'、'.join(missing)}")
            else:
                with st.spinner("正在发送..."):
                    excel_data = generate_review_excel(fields, risks, uploaded.name)
                    subject = f"[ContractLens] 合同审查报告 - {uploaded.name}"
                    body = f"""合同审查报告

文件名: {uploaded.name}
审查时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
甲方: {fields.get('甲方', '—')}
乙方: {fields.get('乙方', '—')}
合同金额: {fields.get('合同金额', '—')}
风险数: {len(risks)} 条

详细内容请查看附件。
---
此邮件由 ContractLens 自动生成
"""
                    ok, msg = send_email_with_excel(
                        email_cfg["server"], email_cfg["port"],
                        email_cfg["sender"], email_cfg["password"],
                        email_cfg["recipient"],
                        subject, body, excel_data,
                        f"ContractLens_{Path(uploaded.name).stem}.xlsx",
                        email_cfg["use_ssl"],
                    )
                    if ok:
                        st.success(f"✅ {msg}")
                    else:
                        st.error(f"❌ {msg}")

    # 存储到 session 供 batch 使用
    st.session_state["last_fields"] = fields
    st.session_state["last_risks"] = risks
    st.session_state["last_summary"] = summary

    # ── 自动保存到审查历史 ──
    saved_id = save_review(uploaded.name, fields, risks, summary, text[:300])
    st.toast(f"✅ 已自动保存到审查历史", icon="💾")


# ═══════════════════════════════════════════════════════════
#  第八部分：批量对比 UI
# ═══════════════════════════════════════════════════════════

def process_single_contract(file_bytes: bytes, filename: str) -> dict[str, Any] | None:
    """处理单个合同，返回汇总信息。"""
    text = extract_text(file_bytes, filename)
    if text.startswith("[错误]") or text.startswith("[提示]"):
        return None

    fields = extract_all_fields(text)
    risks = scan_risks(text)
    summary = generate_summary(text)

    # 自动保存到审查历史
    try:
        save_review(filename, fields, risks, summary, text[:300])
    except Exception:
        pass

    high = sum(1 for r in risks if "高风险" in r["severity"])
    mid = sum(1 for r in risks if "中风险" in r["severity"])
    low = sum(1 for r in risks if "注意" in r["severity"])

    # 计算风险等级
    risk_score = high * 10 + mid * 5 + low * 1
    if risk_score >= 20:
        risk_level = "🔴 重点关注"
    elif risk_score >= 8:
        risk_level = "🟠 需要审阅"
    else:
        risk_level = "🟢 风险较低"

    # 简化金额显示
    amount = fields.get("合同金额", "未提取")
    if amount != "未提取":
        # 提取纯数字部分
        num_match = re.search(r"¥([\d,]+)", amount)
        if num_match:
            amount = "¥" + num_match.group(1)

    return {
        "文件名": filename,
        "甲方": fields.get("甲方", "未提取"),
        "乙方": fields.get("乙方", "未提取"),
        "合同金额": amount,
        "合同期限": fields.get("合同期限", "未提取")[:30],
        "签署日期": fields.get("签署日期", "未提取"),
        "付款节奏": fields.get("付款节奏", "未提取")[:40],
        "风险数": f"🔴{high} 🟠{mid} 🟡{low}",
        "风险分值": risk_score,
        "风险等级": risk_level,
    }


def render_batch_ui():
    """批量上传 + 对比表格界面。"""
    st.header("📊 批量合同对比")

    uploaded_files = st.file_uploader(
        "一次上传多份合同（最多 10 份）",
        type=["pdf", "docx", "doc", "txt"],
        accept_multiple_files=True,
        key="batch_upload",
        help="支持 PDF、DOCX、TXT 格式，最多 10 份",
    )

    if not uploaded_files:
        st.info("👆 上传 2-10 份合同，自动生成横向对比表格")
        st.markdown("""
        #### 对比维度
        - 📋 甲方 / 乙方 — 快速识别交易对手
        - 💰 合同金额 — 金额排序，大额合同优先关注
        - 📅 合同期限 — 是否有自动续约
        - ⚠️ 风险点数 — 🔴🟠🟡 分级统计
        - 🎯 风险等级 — 自动标记「重点关注」合同
        """)
        return

    if len(uploaded_files) > 10:
        st.warning("最多支持同时上传 10 份合同，已自动选取前 10 份")
        uploaded_files = uploaded_files[:10]

    if len(uploaded_files) < 2:
        st.info("请至少上传 2 份合同进行对比")
        return

    # 开始处理
    results: list[dict[str, Any]] = []
    progress = st.progress(0, "正在逐份分析...")
    status_text = st.empty()

    for i, uf in enumerate(uploaded_files):
        status_text.text(f"正在分析 ({i + 1}/{len(uploaded_files)}): {uf.name}")
        try:
            file_bytes = uf.read()
            result = process_single_contract(file_bytes, uf.name)
        except Exception as e:
            result = None
        if result:
            results.append(result)
        else:
            results.append({
                "文件名": f"{uf.name} ⚠️ 解析失败",
                "甲方": "—", "乙方": "—", "合同金额": "—",
                "合同期限": "—", "签署日期": "—", "付款节奏": "—",
                "风险数": "—", "风险分值": 0, "风险等级": "⚠️ 无法分析",
            })
        progress.progress((i + 1) / len(uploaded_files))

    status_text.empty()
    progress.empty()

    if not results:
        st.error("所有文件解析失败")
        return

    # ── 汇总对比表 ──
    st.markdown("### 📊 横向对比表")
    st.caption("按风险分值降序排列，最需关注的合同排在前面")

    df = pd.DataFrame(results)
    df = df.sort_values("风险分值", ascending=False).reset_index(drop=True)

    # 用首列标记风险颜色（避免 styled df 兼容问题）
    colors = []
    for _, row in df.iterrows():
        if "重点关注" in str(row["风险等级"]):
            colors.append("🔴")
        elif "需要审阅" in str(row["风险等级"]):
            colors.append("🟠")
        else:
            colors.append("🟢")
    df.insert(0, "", colors)

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "": st.column_config.TextColumn("", width="small"),
            "文件名": st.column_config.TextColumn("文件名", width="medium"),
            "甲方": st.column_config.TextColumn("甲方", width="medium"),
            "乙方": st.column_config.TextColumn("乙方", width="medium"),
            "合同金额": st.column_config.TextColumn("合同金额", width="small"),
            "合同期限": st.column_config.TextColumn("合同期限", width="medium"),
            "签署日期": st.column_config.TextColumn("签署日期", width="small"),
            "付款节奏": st.column_config.TextColumn("付款节奏", width="medium"),
            "风险数": st.column_config.TextColumn("风险数", width="small"),
            "风险分值": st.column_config.NumberColumn("风险分值", width="small"),
            "风险等级": st.column_config.TextColumn("风险等级", width="small"),
        },
    )

    # ── 统计概览 ──
    st.markdown("---")
    st.markdown("### 📈 批次统计")

    col_s1, col_s2, col_s3, col_s4 = st.columns(4)

    with col_s1:
        st.metric("📄 合同总数", len(results))
    with col_s2:
        total_focus = sum(1 for r in results if "重点关注" in str(r["风险等级"]))
        st.metric("🔴 重点关注", f"{total_focus} 份")
    with col_s3:
        total_review = sum(1 for r in results if "需要审阅" in str(r["风险等级"]))
        st.metric("🟠 需要审阅", f"{total_review} 份")
    with col_s4:
        total_low = sum(1 for r in results if "风险较低" in str(r["风险等级"]))
        st.metric("🟢 风险较低", f"{total_low} 份")

    # ── 金额排行 ──
    st.markdown("### 💰 金额排行 Top 5")
    # 尝试解析金额进行排序
    amount_results = []
    for r in results:
        amt = r["合同金额"]
        sort_key = 0
        if amt not in ("—", "未提取"):
            nums = re.findall(r"[\d,]+", str(amt))
            if nums:
                try:
                    sort_key = float(nums[0].replace(",", ""))
                except ValueError:
                    pass
        amount_results.append({"文件名": r["文件名"], "合同金额": amt, "_sort": sort_key})

    amount_results.sort(key=lambda x: x["_sort"], reverse=True)
    top5 = amount_results[:5]

    if top5 and top5[0]["_sort"] > 0:
        amt_df = pd.DataFrame(top5)
        amt_df = amt_df.drop(columns=["_sort"])
        st.dataframe(amt_df, use_container_width=True, hide_index=True)

    # ── 导出 ──
    st.markdown("---")
    col_x1, col_x2, _ = st.columns([1, 1, 4])

    with col_x1:
        # 导出 Excel 汇总
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="合同对比汇总", index=False)
        excel_buf.seek(0)
        st.download_button(
            label="📥 导出对比 Excel",
            data=excel_buf,
            file_name=f"ContractLens_对比汇总_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_x2:
        # 导出 CSV
        csv_data = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 导出 CSV",
            data=csv_data,
            file_name=f"ContractLens_对比汇总_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # 存储到 session
    st.session_state["batch_results"] = results
    st.session_state["batch_df"] = df


# ═══════════════════════════════════════════════════════════
#  第九部分：审查历史 UI
# ═══════════════════════════════════════════════════════════

def render_history_ui():
    """审查历史管理界面。"""
    st.header("📋 审查历史")

    records = load_all_reviews()

    if not records:
        st.info("📭 暂无审查记录。上传合同并完成审查后，记录会自动保存在这里。")
        return

    # 顶部统计
    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    with col_s1:
        st.metric("📄 总审查数", len(records))
    with col_s2:
        total_high = sum(r.get("high_risk", 0) for r in records)
        st.metric("🔴 高风险合计", total_high)
    with col_s3:
        latest = records[0]
        st.metric("🕐 最近审查", latest.get("timestamp", "")[:10])
    with col_s4:
        st.metric("🗑️ 选中删除", f"{len(st.session_state.get('selected_ids', []))} 条")

    st.markdown("---")

    # 搜索
    search = st.text_input("🔍 搜索文件名", placeholder="输入文件名关键词过滤...")

    # 过滤
    filtered = records
    if search:
        filtered = [r for r in records if search.lower() in r.get("filename", "").lower()]

    if not filtered:
        st.info(f"没有匹配「{search}」的记录")
        return

    # 全选
    all_ids = [r["id"] for r in filtered]
    selected_ids = st.session_state.get("selected_ids", [])

    col_a, col_b = st.columns([1, 5])
    with col_a:
        if st.button("☑️ 全选/取消", use_container_width=True):
            if len(selected_ids) == len(all_ids):
                st.session_state["selected_ids"] = []
            else:
                st.session_state["selected_ids"] = all_ids[:]
            st.rerun()

    # 列表
    for record in filtered:
        rid = record["id"]
        is_selected = rid in selected_ids

        with st.container():
            col_cb, col_info, col_action = st.columns([0.5, 7, 2.5])

            with col_cb:
                checked = st.checkbox("", value=is_selected, key=f"cb_{rid}",
                                      on_change=lambda r=rid: _toggle_select(r))
            with col_info:
                risk_icons = ""
                if record.get("high_risk", 0) > 0:
                    risk_icons += f" 🔴{record['high_risk']}"
                if record.get("mid_risk", 0) > 0:
                    risk_icons += f" 🟠{record['mid_risk']}"
                if record.get("low_risk", 0) > 0:
                    risk_icons += f" 🟡{record['low_risk']}"

                amount = record["fields"].get("合同金额", "—")
                party_a = record["fields"].get("甲方", "—")[:20]

                st.markdown(f"""
                **{record['filename']}**  <small style="color:#94a3b8;">{record['timestamp']}</small>
                <br><small>甲方: {party_a}  |  金额: {amount}  |  风险: {risk_icons if risk_icons else '无'}</small>
                """, unsafe_allow_html=True)

            with col_action:
                col_v, col_d = st.columns(2)
                with col_v:
                    if st.button("📖 查看", key=f"view_{rid}", use_container_width=True):
                        st.session_state["view_record_id"] = rid
                        st.rerun()
                with col_d:
                    if st.button("🗑️", key=f"del_{rid}", use_container_width=True,
                                 help="删除此记录"):
                        delete_review(rid)
                        if rid in st.session_state.get("selected_ids", []):
                            st.session_state["selected_ids"].remove(rid)
                        st.toast("已删除", icon="🗑️")
                        st.rerun()

    # 批量删除
    if selected_ids:
        st.markdown("---")
        col_del, col_mail, _ = st.columns([1, 1, 4])
        with col_del:
            if st.button(f"🗑️ 批量删除选中的 {len(selected_ids)} 条", type="primary", use_container_width=True):
                for rid in selected_ids:
                    delete_review(rid)
                st.session_state["selected_ids"] = []
                st.toast(f"已删除 {len(selected_ids)} 条记录", icon="🗑️")
                st.rerun()
            st.caption("")

    # ── 邮件发送 ──
    with st.expander("📧 发送审查历史到邮箱（点击填写）", expanded=True):
        email_cfg = render_email_settings()

        # 选择发送范围
        send_scope = st.radio(
            "发送范围",
            ["全部记录", f"仅选中 {len(selected_ids)} 条"] if selected_ids else ["全部记录"],
            horizontal=True,
            key="email_scope",
        )

        if st.button("📧 发送 Excel 汇总到邮箱", type="primary"):
            missing = validate_email_settings(email_cfg)
            if missing:
                st.error(f"请填写：{'、'.join(missing)}")
            else:
                with st.spinner("正在生成 Excel 并发送..."):
                    target = records
                    scope_desc = f"共 {len(records)} 条"
                    if "选中" in send_scope and selected_ids:
                        target = [r for r in records if r["id"] in selected_ids]
                        scope_desc = f"选中 {len(target)} 条"

                    excel_data = generate_history_excel(target)
                    subject = f"[ContractLens] 合同审查历史汇总 ({scope_desc})"
                    body = f"""合同审查历史汇总

导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
范围: {scope_desc}

含以下字段：甲方、乙方、合同金额、合同期限、签署日期、生效条件、
付款节奏、违约金、争议解决地、保密期限、风险统计等。

详细内容请查看附件 Excel。
---
此邮件由 ContractLens 自动生成
"""
                    ok, msg = send_email_with_excel(
                        email_cfg["server"], email_cfg["port"],
                        email_cfg["sender"], email_cfg["password"],
                        email_cfg["recipient"],
                        subject, body, excel_data,
                        f"ContractLens_History_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        email_cfg["use_ssl"],
                    )
                    if ok:
                        st.success(f"✅ {msg}")
                    else:
                        st.error(f"❌ {msg}")

    # ── 查看详情弹窗 ──
    view_id = st.session_state.get("view_record_id")
    if view_id:
        detail = load_review(view_id)
        if detail:
            st.markdown("---")
            st.markdown(f"### 📄 {detail['filename']}")
            st.caption(f"审查时间: {detail['timestamp']}")

            # 核心指标
            col_a1, col_b1, col_c1, col_d1 = st.columns(4)
            with col_a1:
                st.metric("甲方", detail["fields"].get("甲方", "—")[:20])
            with col_b1:
                st.metric("乙方", detail["fields"].get("乙方", "—")[:20])
            with col_c1:
                amt = detail["fields"].get("合同金额", "—")
                st.metric("合同金额", amt[:25] if amt != "—" else "—")
            with col_d1:
                st.metric("合同期限", detail["fields"].get("合同期限", "—")[:20])

            # 详细字段
            with st.expander("📋 完整字段", expanded=False):
                for k, v in detail["fields"].items():
                    st.text(f"{k}: {v}")

            # 风险
            risks_detail = detail.get("risks", [])
            if risks_detail:
                with st.expander(f"⚠️ 风险条款（{len(risks_detail)} 条）", expanded=False):
                    for r in risks_detail:
                        st.markdown(f"""
                        <div class="{r.get('css_class', 'risk-low')}">
                            <strong>{r.get('severity', '')} · {r.get('category', '')}</strong>
                            <div class="risk-quote">📌 {r.get('quote', '')}</div>
                        </div>
                        """, unsafe_allow_html=True)

            # 摘要
            with st.expander("📝 合同摘要", expanded=False):
                st.markdown(f'<div class="summary-box">{detail.get("summary", "")}</div>',
                            unsafe_allow_html=True)

            if st.button("✕ 关闭详情"):
                st.session_state.pop("view_record_id", None)
                st.rerun()


def _toggle_select(rid: str):
    """切换选中状态。"""
    selected = st.session_state.get("selected_ids", [])
    if rid in selected:
        selected.remove(rid)
    else:
        selected.append(rid)
    st.session_state["selected_ids"] = selected


# ═══════════════════════════════════════════════════════════
#  第十部分：主入口
# ═══════════════════════════════════════════════════════════

def main():
    """ContractLens 主应用。"""
    _inject_css()

    # 初始化 session state
    if "selected_ids" not in st.session_state:
        st.session_state["selected_ids"] = []
    if "view_record_id" not in st.session_state:
        st.session_state["view_record_id"] = None

    # 侧边栏
    with st.sidebar:
        st.markdown("## 📄 ContractLens")
        st.markdown("*合同速读助手 v1.0*")
        st.markdown("---")

        mode = st.radio(
            "选择模式",
            ["📄 单份审查", "📊 批量对比", "📋 审查历史"],
            label_visibility="collapsed",
        )

        # 只在历史页面显示导出按钮
        if "📋 审查历史" in mode:
            st.markdown("---")
            # 导出全部历史
            all_records = load_all_reviews()
            if all_records:
                json_bytes = export_all_reviews()
                st.download_button(
                    label="📥 导出全部历史 (JSON)",
                    data=json_bytes,
                    file_name=f"ContractLens_History_{datetime.now().strftime('%Y%m%d')}.json",
                    mime="application/json",
                    use_container_width=True,
                )
                # 导入历史
                imported_file = st.file_uploader(
                    "📤 导入历史记录",
                    type=["json"],
                    key="import_history",
                    label_visibility="collapsed",
                )
                if imported_file:
                    count = import_reviews(imported_file.read())
                    if count > 0:
                        st.success(f"✅ 成功导入 {count} 条记录")
                        st.rerun()
                    else:
                        st.warning("⚠️ 没有新记录可导入（可能已存在）")
            st.markdown(f"📊 共 {len(all_records)} 条审查记录")

        st.markdown("---")
        st.markdown("### 🔒 完全离线")
        st.markdown("- 无需联网\n- 无需API密钥\n- 数据不出本地")
        st.markdown("### ⚙️ 技术栈")
        st.markdown("Streamlit + pdfplumber\n+ python-docx + regex\n+ pandas + reportlab")

    # 主界面
    if "📊 批量对比" in mode:
        render_batch_ui()
    elif "📋 审查历史" in mode:
        render_history_ui()
    else:
        render_single_file_ui()


if __name__ == "__main__":
    main()
