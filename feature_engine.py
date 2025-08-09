def extract_features(data: dict):
    """
    data 可来自 Excel/Word 解析；当前给一套示例值兜底，便于演示。
    """
    feats = {
        "bu_count": data.get("bu_count", 3),
        "is_new_market": data.get("is_new_market", True),
        "macro_signals": data.get("macro_signals", True),
        "industry_data": data.get("industry_data", {"cr5": 35, "hhi": 1200, "price_war": True, "switching_cost": "med"}),
        "share_growth": data.get("share_growth", -2),
        "market_growth": data.get("market_growth", 18),
        "internal_data_ready": data.get("internal_data_ready", True),
        "exec_gap": data.get("exec_gap", True),
    }

    # 可选派生
    feats["is_diversified"]   = feats["bu_count"] >= 2
    feats["high_competition"] = (feats["industry_data"].get("hhi", 9999) < 1500) or \
                                (feats["industry_data"].get("cr5", 101) < 40) or \
                                bool(feats["industry_data"].get("price_war"))
    feats["high_growth"]      = feats["market_growth"] is not None and feats["market_growth"] >= 10
    feats["share_problem"]    = feats["share_growth"] is not None and feats["share_growth"] <= 0
    return feats
