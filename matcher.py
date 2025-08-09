def pick_methods(feats, data=None):
    """
    feats 需要包含：
    bu_count, is_new_market, macro_signals,
    industry_data: {cr5, hhi, price_war, switching_cost},
    share_growth, market_growth,
    internal_data_ready, exec_gap
    """
    methods = []

    # A: 宏观触发
    if feats.get("is_new_market") or feats.get("macro_signals"):
        methods.append("PESTEL")

    # B: 行业竞争触发
    ind = feats.get("industry_data", {}) or {}
    cr5 = ind.get("cr5")
    hhi = ind.get("hhi")
    price_war = bool(ind.get("price_war"))
    high_competition = (hhi is not None and hhi < 1500) or (cr5 is not None and cr5 < 40) or price_war
    if high_competition:
        methods.append("FiveForces")  # 名称统一

    # C: 组合分析触发
    if (feats.get("bu_count") or 0) >= 2:
        methods.extend(["BCG", "GE"])

    # D: SWOT 必跑
    methods.append("SWOT")

    # E: 从战略到执行（落地）
    if feats.get("exec_gap") or feats.get("internal_data_ready"):
        methods.append("BLM")

    # 去重并保序
    order = ["PESTEL","FiveForces","BCG","GE","SWOT","BLM"]
    seen, ordered = set(), []
    for m in order:
        if m in methods and m not in seen:
            ordered.append(m); seen.add(m)
    return ordered
