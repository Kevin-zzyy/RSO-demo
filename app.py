import io
import json
import pandas as pd
import streamlit as st
from docx import Document

from report import (
    draw_flow,
    save_markdown_report,
    save_docx_report,
)

st.set_page_config(page_title="战略规划与实施 Demo", layout="wide")
st.title("战略规划与实施知识库 Demo")

# 读取方法论元数据
with open('data/models.json', 'r', encoding='utf-8') as f:
    models_data = json.load(f)

# —— 关键字词典（用于文本提示） ——
METHOD_KEYWORDS = {
    "PESTEL": ["政策", "利率", "关税", "法规", "环保", "技术", "社会", "通胀", "宏观"],
    "FiveForces": ["价格战", "替代品", "进入壁垒", "供应商", "购买者", "份额", "竞争"],
    "BCG": ["份额", "增长率", "现金牛", "明星", "问题", "瘦狗"],
    "GE": ["吸引力", "竞争力", "资源配置"],
    "SWOT": ["优势", "劣势", "机会", "威胁"],
    "BLM": ["OKR", "KPI", "组织", "人才", "里程碑", "RACI", "执行"]
}

# —— 规则引擎 ——

def _display_name(key: str) -> str:
    mapping = {
        "FiveForces": "Porter Five Forces",
        "PESTEL": "PESTEL",
        "BCG": "BCG",
        "GE": "GE",
        "SWOT": "SWOT",
        "BLM": "BLM",
    }
    return mapping.get(key, key)


def pick_methods(feats: dict) -> list:
    methods = []
    if feats.get("is_new_market") or feats.get("macro_signals"):
        methods.append("PESTEL")
    ind = feats.get("industry_data", {}) or {}
    cr5 = ind.get("cr5"); hhi = ind.get("hhi"); price_war = bool(ind.get("price_war"))
    high_comp = (hhi is not None and hhi < 1500) or (cr5 is not None and cr5 < 40) or price_war
    if high_comp:
        methods.append("FiveForces")
    if (feats.get("bu_count") or 0) >= 2:
        methods.extend(["BCG","GE"])
    methods.append("SWOT")
    if feats.get("exec_gap") or feats.get("internal_data_ready"):
        methods.append("BLM")
    order = ["PESTEL","FiveForces","BCG","GE","SWOT","BLM"]
    seen, ordered = set(), []
    for k in order:
        if k in methods and k not in seen:
            ordered.append(_display_name(k)); seen.add(k)
    return ordered


def explain_triggers(feats: dict) -> list:
    ind = feats.get("industry_data", {}) or {}
    reasons = []
    if feats.get("is_new_market") or feats.get("macro_signals"):
        reasons.append("PESTEL：进入新市场或出现宏观信号（is_new_market/macro_signals 为 True）")
    if (ind.get("hhi") is not None and ind.get("hhi") < 1500) or (ind.get("cr5") is not None and ind.get("cr5") < 40) or bool(ind.get("price_war")):
        reasons.append("Porter Five Forces：HHI<1500 或 CR5<40 或存在价格战")
    if (feats.get("bu_count") or 0) >= 2:
        reasons.append("BCG/GE：bu_count≥2，多业务组合需要资源配置建议")
    reasons.append("SWOT：基础综合诊断，默认执行")
    if feats.get("exec_gap") or feats.get("internal_data_ready"):
        reasons.append("BLM：存在执行落差或内部数据可用，需要战略到执行闭环")
    return reasons

# —— 文件解析 ——

def parse_excel(file) -> dict:
    try:
        df = pd.read_excel(file)
        def pick(col, cast=None):
            if col in df.columns and len(df) > 0:
                val = df.iloc[0][col]
                if cast:
                    try:
                        return cast(val)
                    except Exception:
                        return None
                return val
            return None
        return {
            "bu_count": pick("bu_count", int),
            "is_new_market": pick("is_new_market", bool),
            "macro_signals": pick("macro_signals", bool),
            "share_growth": pick("share_growth", float),
            "market_growth": pick("market_growth", float),
            "internal_data_ready": pick("internal_data_ready", bool),
            "exec_gap": pick("exec_gap", bool),
            "industry_data": {
                "cr5": pick("cr5", float),
                "hhi": pick("hhi", float),
                "price_war": pick("price_war", bool),
                "switching_cost": pick("switching_cost", str) or "med",
            }
        }
    except Exception:
        return {}


def parse_word(file) -> str:
    try:
        doc = Document(file)
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    except Exception:
        return ""


def extract_keywords(text: str) -> dict:
    hits = {}
    for m, kws in METHOD_KEYWORDS.items():
        hits[m] = sorted({kw for kw in kws if kw in text})
    return hits

# —— 页面导航 ——
st.sidebar.header("导航")
page = st.sidebar.radio("选择页面", ["概览与说明", "数据上传与特征提取", "方法论库", "生成报告"])

# 概览
if page == "概览与说明":
    st.subheader("支持的方法论板块")
    cols = st.columns(3)
    names = list(models_data.keys())
    for i, name in enumerate(names):
        with cols[i % 3]:
            st.markdown(f"**{name}**\n\n- 简介：{models_data[name].get('简介','')}\n- 应用场景：{models_data[name].get('应用场景','')}")
    with st.expander("📄 上传文件格式说明"):
        st.markdown(
            """
            **Excel（可选）**：推荐列：`bu_count`、`is_new_market`、`macro_signals`、`cr5`、`hhi`、`price_war`、`switching_cost`、`share_growth`、`market_growth`、`internal_data_ready`、`exec_gap`。
            
            **Word（可选）**：自由文本简报；系统会做关键词提示（如“关税/价格战/OKR”等）。
            
            **名词解释**：HHI < 1500 分散，1500–2500 中等，>2500 集中；CR5 = 行业前五份额之和。
            """
        )

# 数据上传与特征提取
elif page == "数据上传与特征提取":
    st.subheader("上传文件与参数填写")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        excel_file = st.file_uploader("上传 Excel（可选）", type=["xlsx", "xls"], key="excel")
        word_file = st.file_uploader("上传 Word（可选）", type=["docx"], key="word")
    with col_u2:
        st.caption("没有文件也可以只用下方表单跑一遍")

    parsed = parse_excel(excel_file) if excel_file else {}
    brief = parse_word(word_file) if word_file else ""

    st.markdown("**关键参数**（可编辑）：")
    c1, c2 = st.columns(2)
    with c1:
        bu_count = st.number_input("业务单元数 bu_count", 1, 100, int(parsed.get("bu_count") or 1))
        is_new_market = st.checkbox("进入新市场/新地区", bool(parsed.get("is_new_market") or False))
        macro_signals = st.checkbox("出现显著宏观变动", bool(parsed.get("macro_signals") or False))
        share_growth = st.number_input("相对份额增长 %", -100.0, 100.0, float(parsed.get("share_growth") or 0.0))
        market_growth = st.number_input("赛道增速 %", -100.0, 500.0, float(parsed.get("market_growth") or 0.0))
    with c2:
        cr5 = st.number_input("CR5 %", 0.0, 100.0, float(parsed.get("industry_data", {}).get("cr5") or 0.0))
        hhi = st.number_input("HHI", 0.0, 10000.0, float(parsed.get("industry_data", {}).get("hhi") or 0.0))
        price_war = st.checkbox("行业存在价格战", bool(parsed.get("industry_data", {}).get("price_war") or False))
        switching_cost = st.selectbox("用户转换成本", ["low", "med", "high"], index=["low","med","high"].index(str(parsed.get("industry_data", {}).get("switching_cost") or "med")))
        internal_data_ready = st.checkbox("内部数据可用（成本/组织/流程）", bool(parsed.get("internal_data_ready") or False))
        exec_gap = st.checkbox("存在战略-执行落差", bool(parsed.get("exec_gap") or False))

    feats = {
        "bu_count": bu_count,
        "is_new_market": is_new_market,
        "macro_signals": macro_signals,
        "share_growth": share_growth,
        "market_growth": market_growth,
        "internal_data_ready": internal_data_ready,
        "exec_gap": exec_gap,
        "industry_data": {"cr5": cr5, "hhi": hhi, "price_war": price_war, "switching_cost": switching_cost},
    }

    kw_hits = {}
    if brief:
        # 简易关键词命中
        for m, kws in METHOD_KEYWORDS.items():
            kw_hits[m] = sorted({kw for kw in kws if kw in brief})
        st.markdown("**从 Word 文本命中的关键词（提示用）**：")
        st.write(kw_hits)

    st.markdown("**提取到的结构化特征**：")
    st.json(feats, expanded=False)

    st.markdown("---")
    run_top = st.button("▶️ 运行分析", type="primary")
    run_side = st.sidebar.button("▶️ 运行分析", key="run_side")
    if run_top or run_side:
        methods = pick_methods(feats)
        st.success("推荐方法论：" + "、".join(methods))
        st.markdown("**触发依据（为什么这么推荐）**：")
        for r in explain_triggers(feats):
            st.write("- ", r)
        st.session_state["feats"] = feats
        st.session_state["methods"] = methods
        st.session_state["kw_hits"] = kw_hits

# 方法论库
elif page == "方法论库":
    st.subheader("方法论一览")
    model_name = st.selectbox("选择战略模型", list(models_data.keys()))
    st.write("**简介：**", models_data[model_name].get("简介", ""))
    st.write("**应用场景：**", models_data[model_name].get("应用场景", ""))
    st.write("**分析步骤：**")
    for step in models_data[model_name].get("分析步骤", []):
        st.write("-", step)

# 生成报告
elif page == "生成报告":
    st.subheader("导出报告")
    st.caption("将上一页的分析结果（方法论/特征/触发依据）汇总为 Markdown 或 Word 报告，并附流程图。")

    use_feats = st.session_state.get("feats", {})
    use_methods = st.session_state.get("methods", [])

    with st.expander("查看当前结果缓存"):
        st.write("方法论：", use_methods)
        st.write("特征：")
        st.json(use_feats, expanded=False)

    fmt = st.radio("选择导出格式", ["Markdown", "Word (docx)"])
    if st.button("生成并下载"):
        triggers = explain_triggers(use_feats) if use_feats else []
        flow_path = draw_flow(use_methods or [])
        if fmt == "Markdown":
            md_path = save_markdown_report(use_methods, use_feats, triggers, models_data)
            with open(md_path, 'rb') as f:
                st.download_button("下载 Markdown 报告", data=f.read(), file_name="战略分析报告.md", mime="text/markdown")
        else:
            docx_path = save_docx_report(use_methods, use_feats, triggers, models_data)
            with open(docx_path, 'rb') as f:
                st.download_button("下载 Word 报告", data=f.read(), file_name="战略分析报告.docx")
        st.image(flow_path, caption="分析流程图", use_container_width=True)
