import os
import json
import pandas as pd


def format_answer_value(ans):
    """Formats answer values into clean text suitable for Excel export."""
    if ans is None:
        return ""
    if isinstance(ans, (int, float, bool)):
        return ans
    if isinstance(ans, str):
        return ans
    if isinstance(ans, list):
        labels = []
        for elem in ans:
            if isinstance(elem, dict):
                if "signature" in elem or (isinstance(elem.get("value"), str) and elem.get("value").startswith("data:image")):
                    labels.append("[Signature]")
                elif "label" in elem and elem["label"]:
                    labels.append(str(elem["label"]))
                elif "value" in elem and elem["value"]:
                    labels.append(str(elem["value"]))
                else:
                    labels.append(json.dumps(elem, ensure_ascii=False))
            else:
                labels.append(str(elem))
        return ", ".join(labels)
    if isinstance(ans, dict):
        if "signature" in ans:
            return "[Signature]"
        if "label" in ans and ans["label"]:
            return str(ans["label"])
        if "value" in ans and ans["value"]:
            return str(ans["value"])
        return json.dumps(ans, ensure_ascii=False)
    return str(ans)


def parse_json_file(filepath):
    """Parses a single respondent JSON file and returns metadata and answers extracted from the 'data' key."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
    except Exception as e:
        print(f"[!] Failed to read {filepath}: {e}")
        return None, None

    main_row = {}

    # 1. Extract top-level metadata keys outside 'data'
    top_meta_keys = [
        "assignment_id",
        "_id",
        "survey_period_id",
        "mode",
        "assignment_status_alias",
        "assignment_error_status_type",
        "code_identity",
        "date_created",
        "date_modified",
        "current_user_fullname",
        "current_user_username",
        "current_user_survey_role_name",
    ]
    for key in top_meta_keys:
        val = raw_data.get(key)
        if val is not None:
            if isinstance(val, list):
                main_row[key] = ", ".join(str(x) for x in val)
            elif isinstance(val, dict):
                main_row[key] = json.dumps(val, ensure_ascii=False)
            else:
                main_row[key] = val

    # Ensure assignment_id is consistently populated
    if "assignment_id" not in main_row or not main_row["assignment_id"]:
        main_row["assignment_id"] = raw_data.get("_id") or raw_data.get("id", "")

    # 2. Extract pre_defined_data predata values if available (fallback/initial metadata)
    pre_data_content = raw_data.get("pre_defined_data")
    if isinstance(pre_data_content, str):
        try:
            pre_obj = json.loads(pre_data_content)
            predata_list = pre_obj.get("predata", [])
            if isinstance(predata_list, list):
                for item in predata_list:
                    if isinstance(item, dict):
                        k = item.get("dataKey")
                        if k and k not in main_row:
                            main_row[k] = format_answer_value(item.get("answer"))
        except Exception:
            pass

    # 3. Extract inner 'data' JSON object
    data_content = raw_data.get("data")
    if isinstance(data_content, str):
        try:
            data_obj = json.loads(data_content)
        except Exception:
            data_obj = {}
    elif isinstance(data_content, dict):
        data_obj = data_content
    else:
        data_obj = {}

    # Include metadata present inside the 'data' object
    meta_keys = [
        "createdAt",
        "createdBy",
        "updatedAt",
        "updatedBy",
        "description",
        "templateVersion",
        "validationVersion",
        "isForceSubmit",
        "dataKey",
    ]
    for meta_key in meta_keys:
        if meta_key in data_obj and data_obj[meta_key] is not None:
            main_row[meta_key] = data_obj[meta_key]

    answers = data_obj.get("answers", [])
    item_rows = []
    items_by_key = {}

    if isinstance(answers, list):
        for item in answers:
            if not isinstance(item, dict):
                continue
            k = item.get("dataKey")
            if not k:
                continue

            val = format_answer_value(item.get("answer"))

            # Handle dynamic keys with '#' hash separators
            if "#" in k:
                parts = k.split("#")
                prefix = parts[0]
                if len(parts) == 3:
                    # Repeating grid/sub-table item: prefix#section_index#item_code
                    sec_idx, item_code = parts[1], parts[2]
                    grid_key = (sec_idx, item_code)
                    if grid_key not in items_by_key:
                        items_by_key[grid_key] = {
                            "assignment_id": main_row.get("assignment_id", ""),
                            "section_index": sec_idx,
                            "item_code": item_code,
                        }
                    items_by_key[grid_key][prefix] = val
                else:
                    # Single section index key: e.g. namakom#1, volume_list#1
                    main_row[k] = val
            else:
                main_row[k] = val

    item_rows = list(items_by_key.values())
    return main_row, item_rows


def extract_json_to_excel(survey_id, period_id, result_dir="result", output_filename="data.xlsx"):
    """Reads all JSON files in result/{survey_id}/{period_id}/json/ and exports extracted 'data' content to Excel."""
    json_dir = os.path.join(result_dir, str(survey_id), str(period_id), "json")
    if not os.path.exists(json_dir):
        print(f"[!] Directory not found: {json_dir}")
        return None

    json_files = [os.path.join(json_dir, f) for f in os.listdir(json_dir) if f.endswith(".json")]
    if not json_files:
        print(f"[!] No JSON files found in {json_dir}")
        return None

    print(f"\n[*] Extracting dynamic data from {len(json_files)} JSON file(s)...")

    all_main_rows = []
    all_item_rows = []

    ordered_main_keys = []
    ordered_item_keys = []

    for filepath in json_files:
        main_row, item_rows = parse_json_file(filepath)
        if main_row is not None:
            all_main_rows.append(main_row)
            for k in main_row.keys():
                if k not in ordered_main_keys:
                    ordered_main_keys.append(k)

            if item_rows:
                for row_dict in item_rows:
                    all_item_rows.append(row_dict)
                    for k in row_dict.keys():
                        if k not in ordered_item_keys:
                            ordered_item_keys.append(k)

    # Base metadata column ordering
    base_main_cols = ["assignment_id"]
    main_cols = base_main_cols + [k for k in ordered_main_keys if k not in base_main_cols]

    df_main = pd.DataFrame(all_main_rows)
    for col in main_cols:
        if col not in df_main.columns:
            df_main[col] = ""
    df_main = df_main[main_cols]

    if all_item_rows:
        base_item_cols = ["assignment_id", "section_index", "item_code"]
        item_cols = base_item_cols + [k for k in ordered_item_keys if k not in base_item_cols]
        df_items = pd.DataFrame(all_item_rows)
        for col in item_cols:
            if col not in df_items.columns:
                df_items[col] = ""
        df_items = df_items[item_cols]
    else:
        df_items = None

    output_dir = os.path.join(result_dir, str(survey_id), str(period_id))
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, output_filename)

    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df_main.to_excel(writer, sheet_name="Data", index=False)
            if df_items is not None:
                df_items.to_excel(writer, sheet_name="Items_Detail", index=False)

        num_items = len(df_items) if df_items is not None else 0
        print(f"[+] Successfully exported dynamic data to: {filepath}")
        print(f"    - Sheet 'Data': {len(df_main)} rows, {len(df_main.columns)} columns")
        if df_items is not None:
            print(f"    - Sheet 'Items_Detail': {num_items} rows, {len(df_items.columns)} columns")

        # Clean up legacy data_detail.xlsx if present
        legacy_path = os.path.join(output_dir, "data_detail.xlsx")
        if os.path.exists(legacy_path):
            try:
                os.remove(legacy_path)
                print(f"[+] Removed legacy file: {legacy_path}")
            except Exception:
                pass

        return filepath
    except Exception as e:
        print(f"[!] Error saving Excel: {e}")
        return None


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3:
        s_id = sys.argv[1]
        p_id = sys.argv[2]
        extract_json_to_excel(s_id, p_id)
    else:
        print("Usage: python extract_excel.py <survey_id> <period_id>")
