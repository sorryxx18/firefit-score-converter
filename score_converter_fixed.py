import pandas as pd
import sys

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
    if item in BIG_BETTER:
        ok = rows[rows["測驗值"] <= value]   # 大值越好
    else:
        ok = rows[rows["測驗值"] >= value]   # 小值越好
    return 0 if ok.empty else ok["得分"].max()

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方式：python score_converter_fixed.py 成績輸入.xlsx 換算表.xlsx 輸出.xlsx")
        sys.exit(1)

    input_file, lookup_file, output_file = sys.argv[1:4]

    df_input  = pd.read_excel(input_file)
    df_lookup = pd.read_excel(lookup_file)

    df_input["年齡層"] = df_input["年齡"].apply(age_to_group)

    BIG_BETTER = {
        "立定跳遠", "後拋擲遠", "折返跑", "菱形槓硬舉",
        "懸吊屈體", "懸吊秒數", "六角槓負重行走"
    }

    ITEM_COL = {
        "立定跳遠":       "立定跳遠(cm)",
        "後拋擲遠":       "後拋擲遠(m)",
        "折返跑":         "折返跑(趟)",
        "菱形槓硬舉":     "菱形槓硬舉(kg)",
        "懸吊屈體":       "懸吊屈體(次)",
        "懸吊秒數":       "懸吊屈體(秒)",
        "六角槓負重行走": "六角槓負重行走(m)",
        "1500跑步":       "1500公尺跑步(秒)"
    }

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

    score_cols = [f"{item}得分" for item in ITEM_COL]
    df_input["總分"] = df_input[score_cols].sum(axis=1)

    df_input.to_excel(output_file, index=False)
    print(f"✅ 換算完成，結果已輸出到 {output_file}")