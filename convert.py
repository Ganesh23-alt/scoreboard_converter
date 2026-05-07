import json
import re
from datetime import datetime
from openpyxl import load_workbook

INPUT_FILE = "scoreboard_test.xlsx"
OUTPUT_FILE = "output.json"
SHEET_NAME = "SCOREBOARD"

# --- Helper Functions ---

def normalize_key(name):
    key = name.lower()
    key = re.sub(r"[^a-z0-9]+", "_", key)
    key = re.sub(r"_+", "_", key)
    key = key.strip("_")
    return key

def clean_value(val):
    if val is None:
        return None
    if isinstance(val, str):
        val = val.strip()
        if val == "":
            return None
    if isinstance(val, float):
        return round(val, 2)
    return val

def get_merged_cell_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for merged in ws.merged_cells.ranges:
        if (
            row >= merged.min_row and row <= merged.max_row and
            col >= merged.min_col and col <= merged.max_col
        ):
            return ws.cell(row=merged.min_row, column=merged.min_col).value
    return None

def build_metric(col_idx, headers, focuses, sources, roles, targets, data_rows, date_col):
    display_name = headers[col_idx]
    if not display_name:
        return None
    metric_key = normalize_key(display_name)
    metric = {
        "metric_key": metric_key,
        "display_name": display_name,
        "focus": focuses.get(col_idx, ""),
        "source": sources.get(col_idx, ""),
        "role": roles.get(col_idx, ""),
        "target": targets.get(col_idx, ""),
        "data": []
    }
    for row in data_rows:
        date_val = clean_value(row[date_col])
        metric_val = clean_value(row.get(col_idx))
        if date_val and metric_val is not None:
            metric["data"].append({
                "date": str(date_val),
                "value": metric_val
            })
    if not metric["data"]:
        return None
    return metric

def extract_metrics(ws):
    HEADER_ROW = 2
    FOCUS_ROW = 3
    SOURCE_ROW = 4
    ROLE_ROW = 5
    TARGET_ROW = 7
    DATA_START_ROW = 8
    headers = {}
    focuses = {}
    sources = {}
    roles = {}
    targets = {}
    date_col = 1
    max_col = ws.max_column
    max_row = ws.max_row
    for col in range(2, max_col + 1):
        headers[col] = clean_value(get_merged_cell_value(ws, HEADER_ROW, col))
        focuses[col] = clean_value(get_merged_cell_value(ws, FOCUS_ROW, col))
        sources[col] = clean_value(get_merged_cell_value(ws, SOURCE_ROW, col))
        roles[col] = clean_value(get_merged_cell_value(ws, ROLE_ROW, col))
        targets[col] = clean_value(get_merged_cell_value(ws, TARGET_ROW, col))
    data_rows = []
    for row in range(DATA_START_ROW, max_row + 1):
        date_val = clean_value(get_merged_cell_value(ws, row, date_col))
        if not date_val:
            continue
        row_data = {date_col: date_val}
        for col in range(2, max_col + 1):
            row_data[col] = clean_value(get_merged_cell_value(ws, row, col))
        data_rows.append(row_data)
    metrics = []
    for col in range(2, max_col + 1):
        metric = build_metric(col, headers, focuses, sources, roles, targets, data_rows, date_col)
        if metric:
            metrics.append(metric)
    return metrics

def build_output(metrics):
    return {
        "sheet_name": SHEET_NAME,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "metrics": metrics
    }

def main():
    wb = load_workbook(INPUT_FILE, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Sheet '{SHEET_NAME}' not found in workbook.")
    ws = wb[SHEET_NAME]
    metrics = extract_metrics(ws)
    output = build_output(metrics)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"Conversion complete. Output written to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
