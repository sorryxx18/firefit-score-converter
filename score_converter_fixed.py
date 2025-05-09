import pandas as pd
import sys

# --- 年齡轉年齡層 (考慮性別) ---
def get_age_group_for_person(row):
    person_name = row.get("姓名", "未知參賽者") # 用於調試輸出
    input_age_group = row.get("年齡層")
    sex = row.get("性別")
    age = row.get("年齡")

    # print(f"DEBUG: 正在處理 {person_name}, 性別: {sex}, 輸入年齡層: {input_age_group}, 輸入年齡: {age}")

    if sex == "女":
        # print(f"DEBUG: {person_name} 為女性，查表年齡層設為 '不分年齡'")
        return "不分年齡"

    if sex == "男":
        if pd.notna(input_age_group) and str(input_age_group).strip() != "":
            age_group_str = str(input_age_group).strip()
            valid_male_age_groups = ["20-29", "30-39", "40-49", "50+"]
            if age_group_str in valid_male_age_groups:
                # print(f"DEBUG: {person_name} (男) 使用手動年齡層: '{age_group_str}'")
                return age_group_str
            else:
                # print(f"DEBUG: {person_name} (男) 手動年齡層 '{age_group_str}' 非標準，嘗試用年齡計算。")
                pass # 繼續嘗試用年齡計算

        if pd.notna(age):
            try:
                age_val = int(float(age))
                if age_val < 30:
                    # print(f"DEBUG: {person_name} (男) 年齡 {age_val} 計算為 '20-29'")
                    return "20-29"
                elif age_val < 40:
                    # print(f"DEBUG: {person_name} (男) 年齡 {age_val} 計算為 '30-39'")
                    return "30-39"
                elif age_val < 50:
                    # print(f"DEBUG: {person_name} (男) 年齡 {age_val} 計算為 '40-49'")
                    return "40-49"
                else:
                    # print(f"DEBUG: {person_name} (男) 年齡 {age_val} 計算為 '50+'")
                    return "50+"
            except ValueError:
                # print(f"DEBUG: {person_name} (男) 年齡 '{age}' 無法轉換為數字，無有效年齡層。")
                return None
        else:
            # print(f"DEBUG: {person_name} (男) 未提供年齡層或有效年齡，無有效年齡層。")
            return None
    
    # print(f"DEBUG: {person_name} 性別 '{sex}' 無法識別或流程未覆蓋，無有效年齡層。")
    return None

# --- 查表得分 (通用邏輯) ---
def get_generic_score(df_lookup, sex, age_group_for_lookup, item_name_for_lookup, value, higher_is_better, person_name="未知"):
    # print(f"DEBUG get_score: 人員:{person_name}, 性別:{sex}, 年齡層:{age_group_for_lookup}, 項目:{item_name_for_lookup}, 測驗值:{value}, 高好?:{higher_is_better}")
    if pd.isna(value) or pd.isna(sex) or pd.isna(age_group_for_lookup) or str(age_group_for_lookup).strip() == "":
        # print(f"DEBUG get_score: 因基礎信息缺失返回0分 (值:{pd.isna(value)}, 性別:{pd.isna(sex)}, 年齡層:{pd.isna(age_group_for_lookup) or str(age_group_for_lookup).strip() == ''})")
        return 0

    try:
        value_numeric = float(value)
    except (ValueError, TypeError):
        # print(f"DEBUG get_score: 測驗值 '{value}' for item '{item_name_for_lookup}' 無法轉為數字，返回0分。")
        return 0

    rows = df_lookup[
        (df_lookup["性別"] == sex) &
        (df_lookup["年齡層"] == age_group_for_lookup) &
        (df_lookup["項目"] == item_name_for_lookup)
    ]

    if rows.empty:
        # print(f"DEBUG get_score: 找不到換算標準 (性別:{sex}, 年齡層:{age_group_for_lookup}, 項目:{item_name_for_lookup})，返回0分。")
        return 0

    valid_rows = rows.dropna(subset=['測驗值']) # 確保換算表的測驗值有效
    if valid_rows.empty:
        # print(f"DEBUG get_score: 篩選到的換算標準中 '測驗值' 列均無效，返回0分。")
        return 0
        
    eligible_scores = pd.DataFrame()
    if higher_is_better:
        eligible_scores = valid_rows[valid_rows["測驗值"] <= value_numeric]
    else:
        eligible_scores = valid_rows[valid_rows["測驗值"] >= value_numeric]

    if eligible_scores.empty:
        # print(f"DEBUG get_score: 根據測驗值 {value_numeric} (高好?:{higher_is_better}) 未找到合格分數段，返回0分。合格標準如下：\n{valid_rows[['得分', '測驗值']]}")
        return 0
    else:
        score = eligible_scores["得分"].max()
        # print(f"DEBUG get_score: 計算得分: {score}")
        return score

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

    df_lkp['測驗值'] = pd.to_numeric(df_lkp['測驗值'], errors='coerce')
    df_in["查表用年齡層"] = df_in.apply(get_age_group_for_person, axis=1)

    SCORING_TYPE = {
        "立定跳遠": True, "後拋擲遠": True, "折返跑": True, "菱形槓硬舉": True,
        "懸吊屈體": True, "懸吊次數": True, "懸吊秒數": True, "負重行走": True,
        "1500跑步": False
    }

    # !!! 請仔細核對這裡的鍵名是否與您「成績輸入.xlsx」的實際欄位名完全一致 !!!
    REGULAR_ITEMS_MAPPING = {
        "立定跳遠(cm)": "立定跳遠",
        "後拋擲遠(m)": "後拋擲遠",
        "折返跑(趟)": "折返跑",
        "菱形槓硬舉(公斤)": "菱形槓硬舉",
        "負重行走": "負重行走", # 如果您的成績輸入表是 "六角槓負重行走", 請修改此處
        "1500公尺跑步(秒)": "1500跑步"
    }
    FLEXION_REPS_INPUT_COL = "懸吊屈體(次)"
    FLEXION_SECS_INPUT_COL = "懸吊屈體(秒)"

    all_score_columns = []

    for input_col_name, lookup_item_name in REGULAR_ITEMS_MAPPING.items():
        score_col_name = f"{lookup_item_name}_得分" # 生成的得分欄位名，如 立定跳遠_得分
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
                SCORING_TYPE[lookup_item_name],
                r.get("姓名", "未知") # 傳遞姓名用於調試
            ), axis=1
        )
    
    # --- 單獨計算懸吊項目得分 ---
    FLEXION_FINAL_SCORE_COL = "懸吊屈體_總得分"
    all_score_columns.append(FLEXION_FINAL_SCORE_COL)

    if FLEXION_REPS_INPUT_COL not in df_in.columns:
        print(f"警告：成績輸入表中找不到欄位 '{FLEXION_REPS_INPUT_COL}'。")
        df_in[FLEXION_REPS_INPUT_COL] = pd.NA
    if FLEXION_SECS_INPUT_COL not in df_in.columns:
        print(f"警告：成績輸入表中找不到欄位 '{FLEXION_SECS_INPUT_COL}'。")
        df_in[FLEXION_SECS_INPUT_COL] = pd.NA

    def calculate_final_flexion_score(row):
        sex = row.get("性別")
        age_group_lookup = row.get("查表用年齡層")
        person_name = row.get("姓名", "未知")
        
        reps_value = row.get(FLEXION_REPS_INPUT_COL)
        secs_value = row.get(FLEXION_SECS_INPUT_COL)

        if sex == "男":
            lookup_item = "懸吊屈體"
            if lookup_item not in SCORING_TYPE:
                 print(f"警告：項目 '{lookup_item}' (男性懸吊) 未在 SCORING_TYPE 中定義。")
                 return 0
            return get_generic_score(df_lkp, sex, age_group_lookup, lookup_item, reps_value, SCORING_TYPE[lookup_item], person_name)
        
        elif sex == "女":
            score_from_reps = 0
            if pd.notna(reps_value):
                try:
                    if float(reps_value) > 0:
                        lookup_item_reps = "懸吊次數"
                        if lookup_item_reps not in SCORING_TYPE:
                            print(f"警告：項目 '{lookup_item_reps}' (女性懸吊次數) 未在 SCORING_TYPE 中定義。")
                        else:
                            score_from_reps = get_generic_score(df_lkp, sex, age_group_lookup, lookup_item_reps, reps_value, SCORING_TYPE[lookup_item_reps], person_name)
                        return score_from_reps
                except (ValueError, TypeError): pass
            
            lookup_item_secs = "懸吊秒數"
            if pd.notna(secs_value):
                if lookup_item_secs not in SCORING_TYPE:
                     print(f"警告：項目 '{lookup_item_secs}' (女性懸吊秒數) 未在 SCORING_TYPE 中定義。")
                     return 0
                return get_generic_score(df_lkp, sex, age_group_lookup, lookup_item_secs, secs_value, SCORING_TYPE[lookup_item_secs], person_name)
            else:
                return 0
        return 0

    df_in[FLEXION_FINAL_SCORE_COL] = df_in.apply(calculate_final_flexion_score, axis=1)

    for col in all_score_columns:
        if col not in df_in.columns:
            df_in[col] = 0
            
    # 確保在加總前，所有得分列都是數值型態，並將NA轉為0
    for col in all_score_columns:
        df_in[col] = pd.to_numeric(df_in[col], errors='coerce').fillna(0)

    df_in["總分"] = df_in[all_score_columns].sum(axis=1)
    
    if "查表用年齡層" in df_in.columns:
        df_in = df_in.drop(columns=["查表用年齡層"])

    try:
        df_in.to_excel(output_file, index=False)
        print(f"✅ 換算完成！結果已輸出到: {output_file}")
    except Exception as e:
        print(f"儲存結果到 Excel 檔案時發生錯誤：{e}")
