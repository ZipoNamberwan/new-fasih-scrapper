import csv
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


def build_roster_schema(template_json):
    """Walks a survey template's components tree and returns:
    - field_roster_path: {dataKey: (ancestor_roster_dataKey, ...)} for every dataKey found,
      ordered outermost -> innermost roster (empty tuple for non-roster/top-level fields)
    - roster_meta: {roster_dataKey: {"label", "fields" (own direct child dataKeys, in
      template order), "ancestor_rosters" (chain of enclosing roster dataKeys)}}
    A "roster" is any component with type == 2 (a repeating group in the Fasih form engine);
    its answers appear as "{childDataKey}#{index}" (or "#{outerIndex}#{innerItemCode}" when
    nested) in the assignment detail's flat answers list.
    """
    field_roster_path = {}
    roster_meta = {}

    def walk(node, roster_chain):
        if isinstance(node, list):
            for item in node:
                walk(item, roster_chain)
            return
        if not isinstance(node, dict):
            return

        dk = node.get("dataKey")
        is_roster = node.get("type") == 2 and dk

        if dk:
            field_roster_path[dk] = tuple(roster_chain)
            if is_roster:
                roster_meta[dk] = {
                    "label": node.get("label") or dk,
                    "fields": [],
                    "ancestor_rosters": tuple(roster_chain),
                }

        next_chain = roster_chain + [dk] if is_roster else roster_chain
        components = node.get("components")
        if components:
            walk(components, next_chain)

    walk(template_json.get("components", []), [])

    for dk, chain in field_roster_path.items():
        if chain and dk not in roster_meta:
            parent_roster = chain[-1]
            if parent_roster in roster_meta:
                roster_meta[parent_roster]["fields"].append(dk)

    return field_roster_path, roster_meta


def load_roster_schema_for_period(survey_id, period_id, result_dir="result"):
    """Loads and merges roster schema (field->roster path, roster metadata) from every
    template JSON saved for this survey period under result/{survey_id}/{period_id}/template/.
    Returns ({}, {}) if no template is available, so callers can fall back to the
    shape-only '#' heuristics instead of failing."""
    template_dir = os.path.join(result_dir, str(survey_id), str(period_id), "template")
    if not os.path.isdir(template_dir):
        return {}, {}

    template_ids = set()
    csv_path = os.path.join(template_dir, "surveyTemplates.csv")
    if os.path.exists(csv_path):
        try:
            with open(csv_path, "r", encoding="utf-8", newline="") as f:
                for row in csv.DictReader(f):
                    tid = row.get("templateId")
                    if tid:
                        template_ids.add(tid)
        except Exception as e:
            print(f"[!] Error reading {csv_path}: {e}")

    if not template_ids:
        template_ids = {
            os.path.splitext(f)[0] for f in os.listdir(template_dir) if f.endswith(".json")
        }

    field_roster_path = {}
    roster_meta = {}
    for tid in template_ids:
        tpath = os.path.join(template_dir, f"{tid}.json")
        if not os.path.exists(tpath):
            continue
        try:
            with open(tpath, "r", encoding="utf-8") as f:
                template_json = json.load(f)
        except Exception as e:
            print(f"[!] Error reading template {tpath}: {e}")
            continue

        fp, rm = build_roster_schema(template_json)
        field_roster_path.update(fp)
        roster_meta.update(rm)

    return field_roster_path, roster_meta


def _index_column_names(depth):
    """Column names for a roster row's repeat-index key, e.g. depth 1 -> ['section_index'],
    depth 2 -> ['section_index', 'item_code'] (matches the existing nested-grid convention)."""
    if depth <= 0:
        return []
    names = ["section_index"]
    if depth >= 2:
        names.append("item_code")
    for extra_level in range(3, depth + 1):
        names.append(f"item_code_{extra_level - 1}")
    return names


_EXCEL_SHEET_ILLEGAL_CHARS = set('[]:*?/\\')


def sanitize_sheet_name(name, used_names):
    """Makes a roster dataKey safe/unique as an Excel sheet name (<=31 chars, no illegal chars)."""
    clean = "".join(c for c in str(name) if c not in _EXCEL_SHEET_ILLEGAL_CHARS).strip()
    if not clean:
        clean = "Sheet"
    clean = clean[:31]

    base = clean
    suffix = 1
    while clean.lower() in used_names:
        suffix_str = f"_{suffix}"
        clean = base[: 31 - len(suffix_str)] + suffix_str
        suffix += 1
    used_names.add(clean.lower())
    return clean


def parse_json_file(filepath, field_roster_path=None, roster_meta=None):
    """Parses a single respondent JSON file. Returns (main_row, roster_rows, legacy_item_rows):
    - main_row: metadata + non-roster answer columns for this respondent
    - roster_rows: {roster_dataKey: [row_dict, ...]} decoded using the survey template; each
      nested roster row is denormalized to also carry its ancestor roster(s)' own field values
    - legacy_item_rows: rows for '#'-keyed answers that don't match any known template field
      (shape-based fallback, used when no template is available or a key is unrecognized)
    """
    field_roster_path = field_roster_path or {}
    roster_meta = roster_meta or {}

    try:
        with open(filepath, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
    except Exception as e:
        print(f"[!] Failed to read {filepath}: {e}")
        return None, None, None

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

    assignment_id = main_row.get("assignment_id", "")

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

    roster_rows_by_key = {}   # roster_dataKey -> {index_tuple: row_dict}
    legacy_items_by_key = {}  # (section_index, item_code) -> row_dict (shape-based fallback)

    if isinstance(answers, list):
        for item in answers:
            if not isinstance(item, dict):
                continue
            k = item.get("dataKey")
            if not k:
                continue

            val = format_answer_value(item.get("answer"))

            if "#" not in k:
                main_row[k] = val
                continue

            parts = k.split("#")
            base = parts[0]
            indices = parts[1:]
            chain = field_roster_path.get(base)

            if chain is not None and len(chain) == len(indices):
                roster_dk = chain[-1]
                row_key = tuple(indices)
                rows_for_roster = roster_rows_by_key.setdefault(roster_dk, {})
                row = rows_for_roster.setdefault(row_key, {"assignment_id": assignment_id})
                for col_name, idx_val in zip(_index_column_names(len(indices)), indices):
                    row[col_name] = idx_val
                row[base] = val
                continue

            # Fallback: shape-based grouping for keys the template doesn't explain
            if len(parts) == 3:
                sec_idx, item_code = parts[1], parts[2]
                grid_key = (sec_idx, item_code)
                if grid_key not in legacy_items_by_key:
                    legacy_items_by_key[grid_key] = {
                        "assignment_id": assignment_id,
                        "section_index": sec_idx,
                        "item_code": item_code,
                    }
                legacy_items_by_key[grid_key][base] = val
            else:
                main_row[k] = val

    # Denormalize: copy each ancestor roster's own field values onto its descendant roster rows
    for roster_dk, rows_by_key in roster_rows_by_key.items():
        ancestor_chain = roster_meta.get(roster_dk, {}).get("ancestor_rosters", ())
        for idx_key, row in rows_by_key.items():
            for depth, ancestor_dk in enumerate(ancestor_chain):
                ancestor_key = idx_key[: depth + 1]
                ancestor_row = roster_rows_by_key.get(ancestor_dk, {}).get(ancestor_key)
                if not ancestor_row:
                    continue
                for field_name in roster_meta.get(ancestor_dk, {}).get("fields", []):
                    if field_name in ancestor_row and field_name not in row:
                        row[field_name] = ancestor_row[field_name]

    roster_rows = {dk: list(rows.values()) for dk, rows in roster_rows_by_key.items()}
    legacy_item_rows = list(legacy_items_by_key.values())

    return main_row, roster_rows, legacy_item_rows


def extract_json_to_excel(survey_id, period_id, result_dir="result", output_filename="data.xlsx"):
    """Reads all JSON files in result/{survey_id}/{period_id}/json/ and exports extracted 'data'
    content to Excel: a 'Data' sheet for non-roster fields, one sheet per repeating-group
    (roster) defined in the survey template, and a legacy 'Items_Detail' fallback sheet for any
    '#'-keyed answers the template doesn't explain (or when no template was downloaded)."""
    json_dir = os.path.join(result_dir, str(survey_id), str(period_id), "json")
    if not os.path.exists(json_dir):
        print(f"[!] Directory not found: {json_dir}")
        return None

    json_files = [os.path.join(json_dir, f) for f in os.listdir(json_dir) if f.endswith(".json")]
    if not json_files:
        print(f"[!] No JSON files found in {json_dir}")
        return None

    field_roster_path, roster_meta = load_roster_schema_for_period(survey_id, period_id, result_dir)
    if roster_meta:
        print(f"[*] Loaded roster schema from template: {len(roster_meta)} roster(s) found.")
    else:
        print("[*] No survey template found; using shape-based '#' heuristics only.")

    print(f"\n[*] Extracting dynamic data from {len(json_files)} JSON file(s)...")

    all_main_rows = []
    ordered_main_keys = []

    all_roster_rows = {}     # roster_dataKey -> list of row dicts
    ordered_roster_cols = {}  # roster_dataKey -> ordered list of column names

    all_legacy_item_rows = []
    ordered_legacy_item_keys = []

    for filepath in json_files:
        main_row, roster_rows, legacy_item_rows = parse_json_file(filepath, field_roster_path, roster_meta)
        if main_row is None:
            continue

        all_main_rows.append(main_row)
        for k in main_row.keys():
            if k not in ordered_main_keys:
                ordered_main_keys.append(k)

        for roster_dk, rows in (roster_rows or {}).items():
            bucket = all_roster_rows.setdefault(roster_dk, [])
            cols = ordered_roster_cols.setdefault(roster_dk, [])
            for row_dict in rows:
                bucket.append(row_dict)
                for k in row_dict.keys():
                    if k not in cols:
                        cols.append(k)

        if legacy_item_rows:
            for row_dict in legacy_item_rows:
                all_legacy_item_rows.append(row_dict)
                for k in row_dict.keys():
                    if k not in ordered_legacy_item_keys:
                        ordered_legacy_item_keys.append(k)

    # Base metadata column ordering
    base_main_cols = ["assignment_id"]
    main_cols = base_main_cols + [k for k in ordered_main_keys if k not in base_main_cols]

    df_main = pd.DataFrame(all_main_rows)
    for col in main_cols:
        if col not in df_main.columns:
            df_main[col] = ""
    df_main = df_main[main_cols]

    roster_dataframes = []  # list of (sheet_name, df)
    used_sheet_names = {"data"}
    for roster_dk, rows in all_roster_rows.items():
        if not rows:
            continue
        base_cols = ["assignment_id"] + [c for c in ("section_index", "item_code") if c in ordered_roster_cols[roster_dk]]
        cols = base_cols + [c for c in ordered_roster_cols[roster_dk] if c not in base_cols]

        df_roster = pd.DataFrame(rows)
        for col in cols:
            if col not in df_roster.columns:
                df_roster[col] = ""
        df_roster = df_roster[cols]

        sheet_name = sanitize_sheet_name(roster_dk, used_sheet_names)
        roster_dataframes.append((sheet_name, df_roster))

    if all_legacy_item_rows:
        base_item_cols = ["assignment_id", "section_index", "item_code"]
        item_cols = base_item_cols + [k for k in ordered_legacy_item_keys if k not in base_item_cols]
        df_items = pd.DataFrame(all_legacy_item_rows)
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
            for sheet_name, df_roster in roster_dataframes:
                df_roster.to_excel(writer, sheet_name=sheet_name, index=False)
            if df_items is not None:
                df_items.to_excel(writer, sheet_name="Items_Detail", index=False)

        print(f"[+] Successfully exported dynamic data to: {filepath}")
        print(f"    - Sheet 'Data': {len(df_main)} rows, {len(df_main.columns)} columns")
        for sheet_name, df_roster in roster_dataframes:
            print(f"    - Sheet '{sheet_name}': {len(df_roster)} rows, {len(df_roster.columns)} columns")
        if df_items is not None:
            print(f"    - Sheet 'Items_Detail': {len(df_items)} rows, {len(df_items.columns)} columns")

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
