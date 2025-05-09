import pandas as pd
import sys

# ── 年齡轉年齡層 ────────────────────────────────────
def age_to_group(age):
    if pd.isna(age):
        return None
    age = int(age)
    if age < 30:
        return "20-29"
    elif age < 40:
        return "30-39"
    elif age < 50:
        return "40-49"
    else:
        return "50+"

# ── 查表得分 ───────────────────────────────────────
def get_score(df_lookup, sex, age_group, item, value):
    if pd.isna(value) or pd.isna(sex) or pd.isna(age_group):
        return 0
    rows = df_lookup[
        (df_lookup["性別"] == sex)
        & (df_lookup["年齡層"] == age_group)
        & (df_lookup["項目"] == item)
    ]
    if rows.empty:
        return 0
    ok = rows[rows["測驗值"] <= value] if item in BIG_BETTER else rows[rows["測驗值"] >= value]
    return 0 if ok.empty else ok["得分"].max()

# ── 主程式 ────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方式：python score_converter_fixed.py  成績輸入頁.xlsx  換算表.xlsx  換算結果.xlsx")
        sys.exit(1)

    input_file, lookup_file, output_file = sys.argv[1:4]

    df_in  = pd.read_excel(input_file)
    df_lkp = pd.read_excel(lookup_file)

    # ➜ 若成績表已有人為填好的「年齡層」，就保留；沒有才用年齡推算
    df_in["年齡層"] = df_in.apply(
        lambda r: r["年齡層"] if pd.notna(r.get("年齡層")) else age_to_group(r["年齡"]),
        axis=1
    )

    # 大值越好
    BIG_BETTER = {
        "立定跳遠", "後拋擲遠", "折返跑",
        "菱形槓硬舉", "懸吊屈體", "懸吊秒數",
        "負重行走"          # ← 與換算表一致
    }

    # 項目 ↔︎ 成績欄
    ITEM_COL = {
        "立定跳遠":   "立定跳遠(cm)",
        "後拋擲遠":   "後拋擲遠(m)",
        "折返跑":     "折返跑(趟)",
        "菱形槓硬舉": "菱形槓硬舉(公斤)",
        "懸吊屈體":   "懸吊屈體(次)",
        "懸吊秒數":   "懸吊屈體(秒)",
        "負重行走":   "負重行走",          # ← 無單位括號
        "1500跑步":   "1500公尺跑步(秒)"
    }

    # 計算各項得分
    for item, col in ITEM_COL.items():
        score_col = f"{item}得分"
        df_in[score_col] = df_in.apply(
            lambda r: get_score(df_lkp, r["性別"], r["年齡層"], item, r.get(col)),
            axis=1
        )

    # 總分
    score_cols = [f"{k}得分" for k in ITEM_COL]
    df_in["總分"] = df_in[score_cols].sum(axis=1)

    df_in.to_excel(output_file, index=False)
    print(f"✅ 換算完成！結果已輸出到: {output_file}")
