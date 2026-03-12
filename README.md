# FreshTake – Document Generator

This repository contains a reusable Python utility for generating rich
**Word (.docx)** documents programmatically, along with the project's
Business Analysis Document.

---

## Repository contents

| Path | Description |
|---|---|
| `Business_Analysis_Document.docx` | Aurum/ORM deployment business analysis document |
| `tools/generate_docx.py` | Reusable DOCX generator library + CLI |
| `requirements.txt` | Python dependencies |

---

## Quick start

### 1 – Install dependencies

```bash
pip install -r requirements.txt
```

### 2 – Generate the built-in example document

```bash
python tools/generate_docx.py --output example.docx
```

Open `example.docx` in Microsoft Word (or LibreOffice) to see a full
demo that includes headings, bullet lists, tables, bar charts, line
charts, and pie charts.

---

## Creating your own document

Import `DocBuilder` and chain the helper methods together:

```python
from tools.generate_docx import DocBuilder

b = DocBuilder(
    title="Project Status Report",
    subtitle="Q2 2026",
    author="Your Name",
    date="March 2026",
)

# Sections & text
b.add_section("1. Summary")
b.add_body("This sprint we completed the login module and started on the dashboard.")
b.add_bullet_list(["Login module ✅", "Dashboard – in progress", "Reporting – not started"])

# Table
b.add_section("2. Task Tracker")
b.add_table(
    headers=["Task", "Owner", "Due Date", "Status"],
    rows=[
        ["Login module",  "Alice", "2026-04-01", "Done"],
        ["Dashboard",     "Bob",   "2026-04-20", "In Progress"],
        ["Reporting",     "Carol", "2026-05-01", "Open"],
    ],
    alt_row_bg="BDD7EE",   # optional alternating row colour
)

# Bar chart
b.add_section("3. Velocity")
b.add_bar_chart(
    title="Sprint Velocity",
    categories=["Sprint 1", "Sprint 2", "Sprint 3"],
    series={"Completed": [34, 41, 38], "Planned": [40, 40, 40]},
    ylabel="Story Points",
    caption="Figure 1 – Sprint velocity.",
)

# Line chart
b.add_line_chart(
    title="Bug Trend",
    categories=["Week 1", "Week 2", "Week 3"],
    series={"Open": [10, 8, 5], "Closed": [2, 6, 9]},
    caption="Figure 2 – Bug count over time.",
)

# Pie chart
b.add_pie_chart(
    title="Effort by Team",
    labels=["Backend", "Frontend", "QA"],
    values=[50, 30, 20],
    caption="Figure 3 – Effort distribution.",
)

b.save("project_status.docx")
```

### Available DocBuilder methods

| Method | What it does |
|---|---|
| `add_section(text, level=1)` | Add a heading (levels 1–4) |
| `add_body(text)` | Add a normal paragraph |
| `add_bullet_list(items, ordered=False)` | Bullet or numbered list |
| `add_note(text, label="Note")` | Indented note paragraph |
| `add_table(headers, rows, ...)` | Styled table with optional coloured header & alternating rows |
| `add_bar_chart(title, categories, series, ...)` | Grouped bar chart embedded as an image |
| `add_line_chart(title, categories, series, ...)` | Line chart embedded as an image |
| `add_pie_chart(title, labels, values, ...)` | Pie chart embedded as an image |
| `add_image(path, width_inches, caption)` | Embed any existing image file |
| `add_page_break()` | Insert a page break |
| `save(output_path)` | Write the finished document to disk |

---

## Reference

- [python-docx documentation](https://python-docx.readthedocs.io/)
- [matplotlib documentation](https://matplotlib.org/stable/index.html)
- [GitHub Markdown syntax guide](https://docs.github.com/github/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax)
