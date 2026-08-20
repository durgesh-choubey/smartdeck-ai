# SmartDeck AI — PULSE AI PowerPoint Generator

Turns a bundled Excel workbook into an updated PowerPoint deck from a
bundled template, in one click.

## How it works

- **`assets/template.pptx`** — the branded deck template (4 slides: Title,
  Key Metrics, Highlights, Revenue Chart), with `{{placeholder}}` markers in
  the text and a native, editable PowerPoint chart on the last slide.
- **`data/sample_data.xlsx`** — the bundled data source (Info / KPIs /
  Highlights / Chart tabs).
- **`src/generator.py`** — reads the Excel data and fills the template's
  placeholders in place. Text is written run-by-run so existing formatting
  (font, size, color) is preserved; the chart is updated via python-pptx's
  native `chart.replace_data()`, so it stays a real, still-editable
  PowerPoint chart — not a flattened image.
- **`app.py`** — the Streamlit interface.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Using it

1. Edit the data in the **Info / KPIs / Highlights / Chart** tabs (or upload
   your own `.xlsx` with the same four tabs — see `data/sample_data.xlsx`
   for the exact column names each tab expects).
2. Click **Update PPT**.
3. Click **Download PPT** to get the `.pptx` file, or turn on
   **Fullscreen Preview** to review the data and the resulting slide content
   at full width first (each `st.dataframe` also has its own native
   full-screen expand icon on hover).

## Customizing the template

Open `assets/template.pptx` in PowerPoint and edit anything you like —
fonts, colors, logo, layout — as long as the `{{marker}}` text runs stay
intact (don't split a marker across multiple text runs, e.g. by
bolding half of it). Re-save over the same file path and the app will use
your updated design on the next run.

To add a new placeholder, add its marker text (e.g. `{{new_field}}`) to a
text box in the template, then add a matching entry to the
`_TEXT_REPLACEMENTS_*` dicts in `src/generator.py`.
