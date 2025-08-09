import networkx as nx
import matplotlib
import matplotlib.pyplot as plt
from datetime import datetime

# 中文字体与负号
matplotlib.rcParams['font.sans-serif'] = ['PingFang SC','Heiti SC','Hiragino Sans GB','Arial Unicode MS','Noto Sans CJK SC','DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

def save_markdown(methods, feats, out_path="demo_output.md"):
    lines = []
    lines.append(f"# 方法论推荐与分析结果\n\n生成时间：{datetime.now().isoformat(timespec='seconds')}\n")
    lines.append("## 已匹配方法论\n" + ", ".join(methods))
    lines.append("\n## 关键特征")
    for k,v in feats.items():
        lines.append(f"- **{k}**: {v}")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_path

def draw_flow(methods, out_path="demo_flow.png"):
    G = nx.DiGraph()
    G.add_edge("客户数据输入", "多维度判断")
    G.add_edge("多维度判断", "方法论匹配引擎")
    for m in methods:
        G.add_edge("方法论匹配引擎", m)

    pos = nx.spring_layout(G, seed=42)
    plt.figure(figsize=(8,6))
    nx.draw(G, pos, with_labels=True, node_size=2200, font_size=10)
    plt.savefig(out_path, dpi=150, bbox_inches='tight')
    plt.close()
    return out_path
