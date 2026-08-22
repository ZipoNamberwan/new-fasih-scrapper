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
    """Parses a single respondent JSON file and returns a flattened dict of metadata and dynamic answers."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
    except Exception as e:
        print(f"[!] Failed to read {filepath}: {e}")
        return None, []

    row = {}

    # Extract base metadata
    row["assignment_id"] = raw_data.get("_id") or raw_data.get("id", "")
    row["survey_period_id"] = raw_data.get("survey_period_id", "")

    mode = raw_data.get("mode")
    row["mode"] = ", ".join(mode) if isinstance(mode, list) else str(mode or "")
    row["assignment_error_status_type"] = raw_data.get("assignment_error_status_type", "")

    # Parse inner 'data'
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

    row["createdAt"] = data_obj.get("createdAt", "")
    row["createdBy"] = data_obj.get("createdBy", "")
    row["updatedBy"] = data_obj.get("updatedBy", "")

    dynamic_keys = []

    # Parse pre_defined_data for initial predata values
    pre_data_content = raw_data.get("pre_defined_data")
    if isinstance(pre_data_content, str):
        try:
            pre_obj = json.loads(pre_data_content)
            predata_list = pre_obj.get("predata", [])
            for item in predata_list:
                k = item.get("dataKey")
                if k and k not in row:
                    val = format_answer_value(item.get("answer"))
                    row[k] = val
                    if k not in dynamic_keys:
                        dynamic_keys.append(k)
        except Exception:
            pass

    # Parse answers array (answers take precedence over predata)
    answers = data_obj.get("answers", [])
    if isinstance(answers, list):
        for item in answers:
            k = item.get("dataKey")
            if k:
                val = format_answer_value(item.get("answer"))
                row[k] = val
                if k not in dynamic_keys:
                    dynamic_keys.append(k)

    return row, dynamic_keys


def extract_json_to_excel(survey_id, period_id, result_dir="result", output_filename="data.xlsx"):
    """Reads all JSON files in result/{survey_id}/{period_id}/json/ and exports them to output_filename (default data.xlsx)."""
    json_dir = os.path.join(result_dir, str(survey_id), str(period_id), "json")
    if not os.path.exists(json_dir):
        print(f"[!] Directory not found: {json_dir}")
        return None

    json_files = [os.path.join(json_dir, f) for f in os.listdir(json_dir) if f.endswith(".json")]
    if not json_files:
        print(f"[!] No JSON files found in {json_dir}")
        return None

    print(f"\n[*] Extracting dynamic detail data from {len(json_files)} JSON file(s)...")

    all_rows = []
    ordered_dynamic_keys = []

    for filepath in json_files:
        row, keys = parse_json_file(filepath)
        if row is not None:
            all_rows.append(row)
            for k in keys:
                if k not in ordered_dynamic_keys:
                    ordered_dynamic_keys.append(k)

    base_cols = [
        "assignment_id",
        "survey_period_id",
        "mode",
        "assignment_error_status_type",
        "createdAt",
        "createdBy",
        "updatedBy",
    ]

    all_detail_cols = base_cols + [k for k in ordered_dynamic_keys if k not in base_cols]

    df_details = pd.DataFrame(all_rows)
    for col in all_detail_cols:
        if col not in df_details.columns:
            df_details[col] = ""

    df_details = df_details[all_detail_cols]

    output_dir = os.path.join(result_dir, str(survey_id), str(period_id))
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, output_filename)

    # If data.xlsx already exists (from save_respondents_to_excel datatable dump), merge them
    if os.path.exists(filepath):
        try:
            print(f"[*] Merging extracted details with existing datatable file: {filepath}")
            df_existing = pd.read_excel(filepath)

            # Normalize ID column name if needed
            if "id" in df_existing.columns and "assignment_id" not in df_existing.columns:
                df_existing.rename(columns={"id": "assignment_id"}, inplace=True)

            if "assignment_id" in df_existing.columns:
                # Drop raw JSON columns from existing datatable if present to keep output clean
                cols_to_drop = [c for c in ["assignment_detail", "detail_data_raw", "detail_predefined_raw"] if c in df_existing.columns]
                if cols_to_drop:
                    df_existing.drop(columns=cols_to_drop, inplace=True)

                # Avoid column duplication on merge
                overlap = [c for c in df_details.columns if c in df_existing.columns and c != "assignment_id"]
                df_existing_clean = df_existing.drop(columns=overlap) if overlap else df_existing

                df_final = pd.merge(df_existing_clean, df_details, on="assignment_id", how="left")
            else:
                df_final = df_details
        except Exception as e:
            print(f"[!] Warning: Could not merge with existing Excel: {e}. Overwriting with detail dataset.")
            df_final = df_details
    else:
        df_final = df_details

    try:
        df_final.to_excel(filepath, index=False)
        print(f"[+] Successfully exported combined data ({len(df_final)} rows, {len(df_final.columns)} columns) to: {filepath}")

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
