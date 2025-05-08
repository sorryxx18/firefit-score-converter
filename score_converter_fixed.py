import pandas as pd
import sys

# ───────────────────────────
# 1. 工具：年齡數字 ➜ 年齡層
# ───────────────────────────
def age_to_group(age: float | int | None) -> str | None:
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

# ───────────────────────────
# 2. 工具：依換算表取得得分
# ───────────────────────────
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

    # 大值越好 vs. 小值越好
    if item in BIG_BETTER:
        ok = rows[rows["測驗值"] <= value]
    else:
        ok = rows[rows["測驗值"] >= value]

    return 0 if ok.empty else ok["得分"].max()

# ───────────────────────────
# 3. 主程式
# ───────────────────────────
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方式：python score_converter_fixed.py  成績輸入.xlsx  換算表.xlsx  輸出.xlsx")
        sys.exit(1)

    input_file, lookup_file, output_file = sys.argv[1:4]

    # 讀檔
    df_input  = pd.read_excel(input_file)
    df_lookup = pd.read_excel(lookup_file)

    # 補年齡層
    df_input["年齡層"] = df_input["年齡"].apply(age_to_group)

    # 大值越好項目
    BIG_BETTER = {
        "立定跳遠",
        "後拋擲遠",
        "折返跑",
        "菱形槓硬舉",
        "懸吊屈體",
        "懸吊秒數",
        "負重行走"          # ← 換算表實際名稱
    }
    # 小值越好：1500跑步（自動走 else 分支）

    # 成績欄位對照（key = 換算表項目；value = 成績輸入檔欄名）
    ITEM_COL = {
        "立定跳遠":   "立定跳遠(cm)",
        "後拋擲遠":   "後拋擲遠(m)",
        "折返跑":     "折返跑(趟)",
        "菱形槓硬舉": "菱形槓硬舉(公斤)",
        "懸吊屈體":   "懸吊屈體(次)",
        "懸吊秒數":   "懸吊屈體(秒)",
        "負重行走":   "負重行走",
        "1500跑步":   "1500公尺跑步(秒)"
    }

    # 計算各項得分
    for item, col in ITEM_COL.items():
        score_col = f"{item}得分"
        df_input[score_col] = df_input.apply(
            lambda row: get_score(
                df_lookup,
                row["性別"],
                row["年齡層"],
                item,
                row.get(col)
            ),
            axis=1
        )

    # 總分
    score_cols = [f"{item}得分" for item in ITEM_COL]
    df_input["總分"] = df_input[score_cols].sum(axis=1)

    # 輸出
    df_input.to_excel(output_file, index=False)
    print(f"✅ 換算完成，結果已輸出到 {output_file}")