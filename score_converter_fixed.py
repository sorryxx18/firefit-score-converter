import pandas as pd
import sys

# --- 年齡轉年齡層 (考慮性別) ---
def get_age_group_for_person(row):
    person_name = row.get("姓名", "未知參賽者")
    input_age_group = row.get("年齡層")
    sex = row.get("性別")
    age = row.get("年齡")

    if sex == "女":
        return "不分年齡"

    if sex == "男":
        if pd.notna(input_age_group) and str(input_age_group).strip() != "":
            age_group_str = str(input_age_group).strip()
            valid_male_age_groups = ["20-29", "30-39", "40-49", "50+"]
            if age_group_str in valid_male_age_groups:
                return age_group_str
        if pd.notna(age):
            try:
                age_val = int(float(age))
                if age_val < 30:
                    return "20-29"
                elif age_val < 40:
                    return "30-39"
                elif age_val < 50:
                    return "40-49"
                else:
                    return "50+"
            except ValueError:
                return None
        else:
            print(f"⚠️ 無法判定年齡層：{person_name} 無年齡與年齡層資料，無法換算分數。")
            return None

    return None

# --- 查表得分 (通用邏輯) ---
def get_generic_score(df_lookup, sex, age_group, item_name, value, higher_is_better, person_name="未知"):
    if pd.isna(value) or pd.isna(sex) or pd.isna(age_group) or str(age_group).strip() == "":
        return 0

    try:
        value_numeric = float(value)
    except (ValueError, TypeError):
        return 0

    rows = df_lookup[
        (df_lookup["性別"] == sex) &
        (df_lookup["年齡層"] == age_group) &
        (df_lookup["項目"] == item_name)
    ]

    if rows.empty:
        return 0

    valid_rows = rows.dropna(subset=['測驗值'])
    if valid_rows.empty:
        return 0

    eligible_scores = valid_rows[valid_rows["測驗值"] <= value_numeric] if higher_is_better else valid_rows[valid_rows["測驗值"] >= value_numeric]

    return eligible_scores["得分"].max() if not eligible_scores.empty else 0

# --- 主程式 ---
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方式：python <script_name>.py 成績輸入.xlsx 換算表.xlsx 換算結果.xlsx")
        sys.exit(1)

    input_file, lookup_file, output_file = sys.argv[1:4]

    try:
        df_in = pd.read_excel(input_file)
        df_lkp = pd.read_excel(lookup_file)
    except Exception as e:
        print(f"讀取 Excel 檔案時發生錯誤：{e}")
        sys.exit(1)

    df_lkp['測驗值'] = pd.to_numeric(df_lkp['測驗值'], errors='coerce')
    df_in["查表用年齡層"] = df_in.apply(get_age_group_for_person, axis=1)

    SCORING_TYPE = {
        "立定跳遠": True, "後拋擲遠": True, "折返跑": True, "菱形槓硬舉": True,
        "懸吊屈體": True, "懸吊次數": True, "懸吊秒數": True, "負重行走": True,
        "1500跑步": False
    }

    REGULAR_ITEMS_MAPPING = {
        "立定跳遠(cm)": "立定跳遠",
        "後拋擲遠(m)": "後拋擲遠",
        "折返跑(趟)": "折返跑",
        "菱形槓硬舉(公斤)": "菱形槓硬舉",
        "負重行走": "負重行走",
        "1500公尺跑步(秒)": "1500跑步"
    }

    FLEXION_REPS_INPUT_COL = "懸吊屈體(次)"
    FLEXION_SECS_INPUT_COL = "懸吊屈體(秒)"
    FLEXION_FINAL_SCORE_COL = "懸吊屈體得分"

    all_score_columns = []

    for input_col, item_name in REGULAR_ITEMS_MAPPING.items():
        score_col = f"{item_name}得分"
        all_score_columns.append(score_col)

        if input_col not in df_in.columns:
            print(f"⚠️ 缺少欄位 '{input_col}'，'{score_col}' 將填入 0")
            df_in[score_col] = 0
            continue

        df_in[score_col] = df_in.apply(
            lambda r: get_generic_score(
                df_lkp, r["性別"], r["查表用年齡層"],
                item_name, r[input_col],
                SCORING_TYPE[item_name],
                r.get("姓名", "未知")
            ), axis=1
        )

    # 懸吊項目處理
    all_score_columns.append(FLEXION_FINAL_SCORE_COL)
    def calculate_flexion_score(row):
        sex = row["性別"]
        age_group = row["查表用年齡層"]
        name = row.get("姓名", "未知")

        reps = row.get(FLEXION_REPS_INPUT_COL)
        secs = row.get(FLEXION_SECS_INPUT_COL)

        if sex == "男":
            return get_generic_score(df_lkp, sex, age_group, "懸吊屈體", reps, SCORING_TYPE["懸吊屈體"], name)
        elif sex == "女":
            if pd.notna(reps):
                try:
                    if float(reps) > 0:
                        return get_generic_score(df_lkp, sex, age_group, "懸吊次數", reps, SCORING_TYPE["懸吊次數"], name)
                except:
                    pass
            if pd.notna(secs):
                return get_generic_score(df_lkp, sex, age_group, "懸吊秒數", secs, SCORING_TYPE["懸吊秒數"], name)
            return 0
        return 0

    df_in[FLEXION_FINAL_SCORE_COL] = df_in.apply(calculate_flexion_score, axis=1)

    # 加總總分
    for col in all_score_columns:
        df_in[col] = pd.to_numeric(df_in[col], errors='coerce').fillna(0)
    df_in["總分"] = df_in[all_score_columns].sum(axis=1)

    if "查表用年齡層" in df_in.columns:
        df_in.drop(columns=["查表用年齡層"], inplace=True)

    try:
        df_in.to_excel(output_file, index=False)
        print(f"✅ 換算完成，已輸出至：{output_file}")
    except Exception as e:
        print(f"❌ 儲存 Excel 發生錯誤：{e}")
