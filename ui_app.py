import streamlit as st
import pandas as pd
from docx import Document
from feature_engine import extract_features
from matcher import pick_methods
from report import save_markdown, draw_flow

st.set_page_config(page_title="RSO Demo", layout="centered")
st.title("RSO 知识库 Demo")
st.caption("上传 Excel / Word（可选），或直接填写参数；点击“运行分析”获取推荐方法论、报告与流程图。")

# ===== 文件上传（可选） =====
excel_file = st.file_uploader("上传 Excel（供应链数据，可选）", type=["xlsx", "xls"])
word_file  = st.file_uploader("上传 Word（业务/宏观简报，可选）", type=["docx"])

# ===== 参数表单（便于演示与兜底） =====
with st.form("params"):
    st.write("参数（没有文件也可直接用这些参数跑）：")
    col1, col2 = st.columns(2)
    with col1:
        bu_count = st.number_input("业务单元数（bu_count）", min_value=1, max_value=50, value=3)
        is_new_market = st.checkbox("进入新市场/新地区？", value=True)
        macro_signals = st.checkbox("出现显著宏观变动？", value=True)
        share_growth = st.number_input("相对份额增长 %（可为负）", min_value=-100, max_value=100, value=-2)
        market_growth = st.number_input("赛道增速 %（YoY）", min_value=-100, max_value=500, value=18)
    with col2:
        price_war = st.checkbox("行业价格战？", value=True)
        cr5 = st.number_input("CR5（%）", min_value=0, max_value=100, value=35)
        hhi = st.number_input("HHI", min_value=0, max_value=10000, value=1200)
        internal_data_ready = st.checkbox("内部成本/组织/流程数据可用？", value=True)
        exec_gap = st.checkbox("存在战略-执行落差？", value=True)

    submitted = st.form_submit_button("运行分析")

# ===== 解析上传内容（可选，不强制） =====
uploaded_info = {}
if excel_file is not None:
    try:
        df = pd.read_excel(excel_file)
        uploaded_info["excel_rows"] = df.shape[0]
        uploaded_info["excel_cols"] = df.shape[1]
        st.caption(f"✓ 已读取 Excel：{df.shape[0]} 行 × {df.shape[1]} 列")
    except Exception as e:
        st.warning(f"读取 Excel 失败：{e}")

brief_text = ""
if word_file is not None:
    try:
        doc = Document(word_file)
        brief_text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
        st.caption(f"✓ 已读取 Word：约 {len(brief_text)} 字")
    except Exception as e:
        st.warning(f"读取 Word 失败：{e}")

# ===== 组装引擎输入（与 feature_engine 对接） =====
data = {
    "bu_count": bu_count,
    "is_new_market": is_new_market,
    "macro_signals": macro_signals,
    "share_growth": share_growth,
    "market_growth": market_growth,
    "internal_data_ready": internal_data_ready,
    "exec_gap": exec_gap,
    "industry_data": {
        "cr5": cr5,
        "hhi": hhi,
        "price_war": price_war,
        "switching_cost": "med",
    },
    # 下面两个字段目前仅占位，后续需要时可在 extract_features 里利用
    "excel_meta": uploaded_info,
    "brief": brief_text,
}

# ===== 运行分析 =====
if submitted:
    st.write("🧪 提取特征…")
    feats = extract_features(data)

    st.write("🧠 匹配方法论…")
    methods = pick_methods(feats, data)
    st.success("推荐方法论： " + "、".join(methods))

    st.write("📝 生成报告与流程图…")
    md_path = save_markdown(methods, feats, "demo_output.md")
    png_path = draw_flow(methods, "demo_flow.png")

    # 下载与展示
    with open(md_path, "rb") as f:
        st.download_button("下载报告（Markdown）", data=f.read(), file_name="demo_output.md", mime="text/markdown")
    st.image(png_path, caption="分析流程图", use_container_width=True)

    # 可选显示特征与输入，便于面试/汇报解释
    with st.expander("查看引擎输入与特征（用于可解释性）"):
        st.write("data：", data)
        st.write("features：", feats)
