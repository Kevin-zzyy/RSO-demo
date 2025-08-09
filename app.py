from feature_engine import extract_features
from matcher import pick_methods
from report import save_markdown, draw_flow

def load_data():
    print("[1/4] 读取客户数据（占位，可接 Excel/Word）")
    return {}, None  # 先用空数据，extract_features 会给示例兜底

def make_report(methods, feats):
    print("[4/4] 生成报告与流程图")
    md = save_markdown(methods, feats, "demo_output.md")
    png = draw_flow(methods, "demo_flow.png")
    print(f"已生成：{md}")
    print(f"流程图：{png}")

if __name__ == "__main__":
    data, _ = load_data()
    print("[2/4] 提取特征")
    feats = extract_features(data)
    print("[3/4] 方法论匹配")
    methods = pick_methods(feats, data)
    make_report(methods, feats)
