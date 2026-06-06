---
{
  "id": "file_0y639z7g",
  "filetype": "document",
  "filename": "google-sheets-placement-sorter",
  "created_at": "2026-06-06T00:52:28.516Z",
  "updated_at": "2026-06-06T00:53:22.370Z",
  "meta": {
    "location": "/",
    "tags": [],
    "categories": [],
    "description": "",
    "source": "markdown"
  }
}
---
# Google Sheets Placement Preference Sorter

This script sorts first-year medical school programme rows from most preferred to least preferred based on the placement preferences you enter.

## Install

1. Open the Google Sheet.
2. If the file still says `.XLSX` beside the name, click `File > Save as Google Sheets` first.
3. Go to `Extensions > Apps Script`.
4. Paste the contents of `scripts/placement-preference-sorter.gs` into the Apps Script editor.
5. Save the project.
6. Reload the Google Sheet.
7. Use the new `Placement Sorter` menu.

## Use

1. Click `Placement Sorter > Create / Refresh Setup Sheets`.
2. Open the `Sorter Settings` tab and check the settings.
3. Open the `Preferences` tab.
4. Reorder column B so your favourite placement is at the top and your least favourite is at the bottom.
5. Click `Placement Sorter > Sort Programmes by Preferences`.

## Non-coder settings

The `Sorter Settings` tab lets you change the script without editing code.

| Setting | What it means | Default |
| :--- | :--- | :--- |
| Target sheet name | The tab containing the programme table. Leave blank to use the first normal sheet. | blank |
| First data row | First row containing a programme. | `3` |
| Rank column | Column where rank numbers should be written. | `E` |
| Programme code column | Column containing codes like `FY1-26-001`. | `F` |
| First placement column | First quarter placement column. | `G` |
| Number of placement columns | Number of placement columns to score. | `4` |
| Rank header cell | Cell containing the rank column heading. | `E1` |
| Preference weighting strength | How strongly top preferences are prioritised. | `2` |
| Preserve colours and text styling | Whether colours/fonts move with each programme row. | `Yes` |

## Weighting guide

`Preference weighting strength` controls how much the top preferences dominate the ranking.

- `1` = gentle weighting. Balanced rows can compete with rows containing one top choice.
- `2` = balanced weighting. Good default.
- `3` = strong weighting. Top choices have a much bigger effect.
- `4` = very strong weighting. Best if you want top choices to dominate heavily.

## How it ranks programmes

Each programme row gets a weighted score from the four placement columns. Top preferences receive more points than lower preferences, so programmes containing your favourite placements are more likely to appear near the top.

For example, if your preferences are:

```text
Medicine
Surgery
Cardiology
Urology
```

Then `Medicine` receives the strongest score, `Surgery` receives the next strongest score, `Cardiology` receives the next, and so on. The programme row with the highest total weighted score is ranked highest.

Ties are broken by the best individual placement in the row, then the second best, then the programme code.

## Important note

Run this on a copy first. Sorting changes the order of rows in the sheet, though it preserves the row values and the main colours/fonts across the ranked table when `Preserve colours and text styling` is set to `Yes`.