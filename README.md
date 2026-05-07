# Scoreboard Excel → JSON Converter

## How to Run

1. Install Python 3.11+ and dependencies:
   ```sh
   pip install -r requirements.txt
   ```
2. Place `Scoreboard Test.xlsx` in this directory.
3. Run:
   ```sh
   python convert.py
   ```

## JSON Output Shape

The script produces a JSON file with this structure:

```
{
  "sheet_name": "SCOREBOARD",
  "generated_at": "2026-05-07T12:00:00",
  "metrics": [
    {
      "metric_key": "total_revenue_all_services",
      "display_name": "Total Revenue - All Services",
      "focus": "Financial",
      "source": "EMR",
      "role": "J",
      "target": "100%",
      "data": [
        { "date": "2026-02-16", "value": 40454.28 }
      ]
    }
  ]
}
```

## Design & Trade-offs

- **Merged cells**: Flattened so values are always available for each metric.
- **Formulas**: Only computed values are extracted (using `data_only=True`).
- **Spacer columns**: Omitted from output.
- **Formatting**: All visual formatting is ignored.
- **Normalization**: Metric keys are normalized for developer/LLM use.

## With More Time

- Add CLI/config for flexible input/output
- More robust handling of edge cases
- Unit tests
- Support for multiple sheets

## Recent Improvements

The converter was refactored to improve performance and data usability:

- Optimized merged cell handling using precomputed lookup map
- Improved JSON structure for analytics-friendly consumption
- Removed redundant Excel-style row processing
- Enhanced date parsing and normalization
- Reduced runtime complexity for large spreadsheets
- Improved maintainability with cleaner separation of logic

This version is designed to support downstream use cases such as dashboards, APIs, and LLM-based querying.