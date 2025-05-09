import pandas as pd
import sys

# --- 年齡轉年齡層 (考慮性別) ---
def get_age_group_for_person(row):
    input_age_group = row.get("年齡層") # .get() 避免原始成績表無此欄位時出錯
    sex = row.get("性別")
    age = row.get("年齡")

    if sex == "女":
        return "不分年齡"

    if sex == "男":
        if pd.notna(input_age_group) and str(input_age_group).strip() != "": # strip() 去除可能的空白字串
            # 確保男性手動填寫的年齡層是換算表接受的格式
            valid_male_age_groups = ["20-29", "30-39", "40-49", "50+"]
            if str(input_age_group).strip() in valid_male_age_groups:
                return str(input_age_group).strip()
            # else:
                # print(f"警告: 參賽者 {row.get('姓名', '未知')} 的手動年齡層 '{input_age_group}' 非標準格式，將嘗試用年齡計算。")
                # Fall through to age calculation if manually entered group is not standard

        if pd.notna(age):
            try:
                age_val = int(float(age)) # 先轉 float 允許小數點，再轉 int
                if age_val < 30:
                    return "20-29"
                elif age_val < 40:
                    return "30-39"
                elif age_val < 50:
                    return "40-49"
                else:
                    return "50+"
            except ValueError:
                # print(f"警告: 參賽者 {row.get('姓名', '未知')} 的年齡 '{age}' 無法轉換為數字。")
                return None # 年齡無效
        else:
            # print(f"警告: 男性參賽者 {row.get('姓名', '未知')} 未提供年齡層或有效年齡。")
            return None # 男性無年齡也無年齡層
    
    # print(f"警告: 參賽者 {row.get('姓名', '未知')} 性別 '{sex}' 無法識別或不完整。")
    return None # 其他性別或情況

# --- 查表得分 (通用邏輯) ---
def get_generic_score(df_lookup, sex, age_group_for_lookup, item_name_for_lookup, value, higher_is_better):
    if pd.isna(value) or pd.isna(sex) or pd.isna(age_group_for_lookup) or str(age_group_for_lookup).strip() == "":
        return 0

    try:
        value_numeric = float(value)
    except (ValueError, TypeError):
        # print(f"Debug: 測驗值 '{value}' for item '{item_name_for_lookup}' 無法轉為數字。")
        return 0

    # 確保換算表中的 '測驗值' 是數字 (這步理論上在讀取df_lkp後做一次即可)
    # df_lookup['測驗值'] = pd.to_numeric(df_lookup['測驗值'], errors='coerce') # 移到主程式

    rows = df_lookup[
        (df_lookup["性別"] == sex) &
        (df_lookup["年齡層"] == age_group_for_lookup) &
        (df_lookup["項目"] == item_name_for_lookup)
    ]

    if rows.empty:
        # print(f"Debug: 找不到換算標準 - 性別:{sex}, 年齡層:{age_group_for_lookup}, 項目:{item_name_for_lookup}")
        return 0

    eligible_scores = pd.DataFrame()
    # 確保 rows['測驗值'] 沒有 NaN (to_numeric errors='coerce' 會產生NaN)
    valid_rows = rows.dropna(subset=['測驗值'])

    if higher_is_better:
        eligible_scores = valid_rows[valid_rows["測驗值"] <= value_numeric]
    else:
        eligible_scores = valid_rows[valid_rows["測驗值"] >= value_numeric]

    if eligible_scores.empty:
        return 0
    else:
        return eligible_scores["得分"].max()

# --- 主程式 ---
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("使用方式：python <script_name>.py 成績輸入.xlsx 換算表.xlsx 換算結果.xlsx")
        sys.exit(1)

    input_file, lookup_file, output_file = sys.argv[1:4]

    try:
        df_in = pd.read_excel(input_file)
        df_lkp = pd.read_excel(lookup_file)
    except FileNotFoundError as e:
        print(f"錯誤：找不到檔案 {e.filename}")
        sys.exit(1)
    except Exception as e:
        print(f"讀取 Excel 檔案時發生錯誤：{e}")
        sys.exit(1)

    # --- 預處理和定義 ---
    # 確保換算表 '測驗值' 為數字
    df_lkp['測驗值'] = pd.to_numeric(df_lkp['測驗值'], errors='coerce')


    # 1. 統一處理 df_in 中的 '年齡層'，供後續查表使用
    df_in["查表用年齡層"] = df_in.apply(get_age_group_for_person, axis=1)

    # 2. 定義哪些項目是數值越高越好 (鍵: 換算表中的「項目」名稱)
    SCORING_TYPE = {
        "立定跳遠": True, "後拋擲遠": True, "折返跑": True, "菱形槓硬舉": True,
        "懸吊屈體": True, "懸吊次數": True, "懸吊秒數": True, "負重行走": True,
        "1500跑步": False
    }

    # 3. 項目名稱映射 (鍵: 成績輸入表欄位名, 值: 換算表「項目」名)
    REGULAR_ITEMS_MAPPING = {
        "立定跳遠(cm)": "立定跳遠", "後拋擲遠(m)": "後拋擲遠",
        "折返跑(趟)": "折返跑", "菱形槓硬舉(公斤)": "菱形槓硬舉",
        "負重行走": "負重行走", "1500公尺跑步(秒)": "1500跑步"
    }
    # 懸吊項目在成績輸入表中的欄位名
    FLEXION_REPS_INPUT_COL = "懸吊屈體(次)"
    FLEXION_SECS_INPUT_COL = "懸吊屈體(秒)"


    all_score_columns = []

    # --- 計算常規項目得分 ---
    for input_col_name, lookup_item_name in REGULAR_ITEMS_MAPPING.items():
        score_col_name = f"{lookup_item_name}_得分"
        all_score_columns.append(score_col_name)

        if input_col_name not in df_in.columns:
            print(f"警告：成績輸入表中找不到欄位 '{input_col_name}'，'{score_col_name}' 將全為0。")
            df_in[score_col_name] = 0
            continue
            
        if lookup_item_name not in SCORING_TYPE:
            print(f"警告：項目 '{lookup_item_name}' 未在 SCORING_TYPE 中定義計分方式，'{score_col_name}' 將全為0。")
            df_in[score_col_name] = 0
            continue

        df_in[score_col_name] = df_in.apply(
            lambda r: get_generic_score(
                df_lkp, r.get("性別"), r.get("查表用年齡層"),
                lookup_item_name, r.get(input_col_name),
                SCORING_TYPE[lookup_item_name]
            ), axis=1
        )

    # --- 單獨計算懸吊項目得分 ---
    FLEXION_FINAL_SCORE_COL = "懸吊屈體_總得分" # 統一的懸吊得分欄位名
    all_score_columns.append(FLEXION_FINAL_SCORE_COL)

    # 檢查懸吊相關欄位是否存在於成績輸入表
    if FLEXION_REPS_INPUT_COL not in df_in.columns:
        print(f"警告：成績輸入表中找不到欄位 '{FLEXION_REPS_INPUT_COL}'，懸吊(次)成績可能無法計算。")
        df_in[FLEXION_REPS_INPUT_COL] = pd.NA # 補上缺失欄位以防 apply 出錯
    if FLEXION_SECS_INPUT_COL not in df_in.columns:
        print(f"警告：成績輸入表中找不到欄位 '{FLEXION_SECS_INPUT_COL}'，懸吊(秒)成績可能無法計算。")
        df_in[FLEXION_SECS_INPUT_COL] = pd.NA # 補上缺失欄位

    def calculate_final_flexion_score(row):
        sex = row.get("性別")
        age_group_lookup = row.get("查表用年齡層")
        
        reps_value = row.get(FLEXION_REPS_INPUT_COL)
        secs_value = row.get(FLEXION_SECS_INPUT_COL)

        if sex == "男":
            lookup_item = "懸吊屈體" # 換算表項目名
            if lookup_item not in SCORING_TYPE:
                 print(f"警告：項目 '{lookup_item}' (男性懸吊) 未在 SCORING_TYPE 中定義，得分將為0。")
                 return 0
            return get_generic_score(df_lkp, sex, age_group_lookup, lookup_item, reps_value, SCORING_TYPE[lookup_item])
        
        elif sex == "女":
            score_from_reps = 0
            # 優先用次數，但次數需 > 0
            if pd.notna(reps_value):
                try:
                    if float(reps_value) > 0:
                        lookup_item_reps = "懸吊次數" # 換算表項目名
                        if lookup_item_reps not in SCORING_TYPE:
                            print(f"警告：項目 '{lookup_item_reps}' (女性懸吊次數) 未在 SCORING_TYPE 中定義，得分將為0。")
                        else:
                            score_from_reps = get_generic_score(df_lkp, sex, age_group_lookup, lookup_item_reps, reps_value, SCORING_TYPE[lookup_item_reps])
                        return score_from_reps # 次數>0，以此為準
                except (ValueError, TypeError):
                    pass # reps_value 不是數字或無法比較，則忽略，繼續看秒數
            
            # 若次數為0, NaN, 或非數字, 則看秒數
            lookup_item_secs = "懸吊秒數" # 換算表項目名
            if pd.notna(secs_value):
                if lookup_item_secs not in SCORING_TYPE:
                     print(f"警告：項目 '{lookup_item_secs}' (女性懸吊秒數) 未在 SCORING_TYPE 中定義，得分將為0。")
                     return 0
                return get_generic_score(df_lkp, sex, age_group_lookup, lookup_item_secs, secs_value, SCORING_TYPE[lookup_item_secs])
            else:
                return 0 # 次數和秒數都無有效成績或不適用
        return 0

    df_in[FLEXION_FINAL_SCORE_COL] = df_in.apply(calculate_final_flexion_score, axis=1)

    # --- 計算總分 ---
    for col in all_score_columns:
        if col not in df_in.columns:
            df_in[col] = 0 # 確保所有計分欄位存在
            
    df_in["總分"] = df_in[all_score_columns].sum(axis=1, skipna=True) # skipna 以防萬一有NA
    
    if "查表用年齡層" in df_in.columns:
        df_in = df_in.drop(columns=["查表用年齡層"])

    # --- 輸出結果 ---
    try:
        df_in.to_excel(output_file, index=False)
        print(f"✅ 換算完成！結果已輸出到: {output_file}")
    except Exception as e:
        print(f"儲存結果到 Excel 檔案時發生錯誤：{e}")
