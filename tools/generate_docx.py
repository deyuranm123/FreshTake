"""
generate_docx.py
================
Reusable helper library + CLI for generating rich Word (.docx) documents
programmatically using python-docx and matplotlib.

Quick-start
-----------
    python tools/generate_docx.py --example            # write example.docx
    python tools/generate_docx.py --output my_doc.docx # same, custom filename

Using as a library
------------------
    from tools.generate_docx import DocBuilder

    b = DocBuilder("My Report Title", subtitle="Draft v1.0")
    b.add_section("Introduction")
    b.add_body("This is the opening paragraph of the report.")
    b.add_bullet_list(["Point one", "Point two", "Point three"])
    b.add_table(
        headers=["Item", "Owner", "Due Date", "Status"],
        rows=[
            ["Set up CI", "Alice", "2026-04-01", "Open"],
            ["Write tests", "Bob",   "2026-04-15", "In Progress"],
        ],
    )
    b.add_bar_chart(
        title="Sprint Velocity",
        categories=["Sprint 1", "Sprint 2", "Sprint 3"],
        series={"Points Completed": [34, 41, 38]},
    )
    b.save("my_report.docx")
"""

from __future__ import annotations

import argparse
import io
from typing import Any

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def _set_cell_bg(cell: Any, hex_color: str) -> None:
    """Fill a table cell with a solid background colour (hex, e.g. '1F3864')."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _rgb(hex_color: str) -> RGBColor:
    """Convert a 6-char hex string to an RGBColor."""
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------------------
# DocBuilder – main public API
# ---------------------------------------------------------------------------

class DocBuilder:
    """
    Fluent helper for building Word documents.

    Parameters
    ----------
    title:
        Main document title (shown on the cover page).
    subtitle:
        Optional subtitle shown below the main title.
    author:
        Author name added to cover-page metadata line.
    date:
        Date string added to cover-page metadata line.
    accent_hex:
        6-char hex colour used for headings and table headers (default dark-blue).
    """

    # Default palette
    _ACCENT      = "1F3864"   # dark navy
    _ACCENT2     = "4472C4"   # medium blue
    _HEADER_FG   = "FFFFFF"
    _TABLE_STYLE = "Table Grid"

    def __init__(
        self,
        title: str,
        subtitle: str = "",
        author: str = "",
        date: str = "",
        accent_hex: str = "1F3864",
    ) -> None:
        self._accent = accent_hex.lstrip("#")
        self._doc = Document()

        # Page margins
        for section in self._doc.sections:
            section.top_margin    = Cm(2.0)
            section.bottom_margin = Cm(2.0)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        # Base font
        self._doc.styles["Normal"].font.name = "Calibri"
        self._doc.styles["Normal"].font.size = Pt(11)

        # Cover page
        self._build_cover(title, subtitle, author, date)
        self._doc.add_page_break()

    # ------------------------------------------------------------------
    # Cover page
    # ------------------------------------------------------------------

    def _build_cover(
        self, title: str, subtitle: str, author: str, date: str
    ) -> None:
        doc = self._doc
        doc.add_paragraph()

        tp = doc.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = tp.add_run(title)
        run.bold = True
        run.font.size = Pt(26)
        run.font.color.rgb = _rgb(self._accent)

        if subtitle:
            sp = doc.add_paragraph()
            sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            sr = sp.add_run(subtitle)
            sr.font.size = Pt(14)
            sr.font.color.rgb = _rgb("4472C4")

        doc.add_paragraph()
        meta = {k: v for k, v in [("Author", author), ("Date", date)] if v}
        for key, val in meta.items():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p.add_run(f"{key}:  ")
            r1.bold = True
            r1.font.color.rgb = _rgb(self._accent)
            p.add_run(val)

    # ------------------------------------------------------------------
    # Sections / headings
    # ------------------------------------------------------------------

    def add_section(self, text: str, level: int = 1) -> "DocBuilder":
        """Add a heading at *level* (1–4)."""
        h = self._doc.add_heading(text, level=level)
        for run in h.runs:
            run.font.color.rgb = _rgb(self._accent)
        return self

    def add_page_break(self) -> "DocBuilder":
        self._doc.add_page_break()
        return self

    # ------------------------------------------------------------------
    # Text blocks
    # ------------------------------------------------------------------

    def add_body(self, text: str) -> "DocBuilder":
        """Add a normal body paragraph."""
        self._doc.add_paragraph(text)
        return self

    def add_bullet_list(
        self, items: list[str], ordered: bool = False
    ) -> "DocBuilder":
        """Add a bullet (or numbered) list."""
        style = "List Number" if ordered else "List Bullet"
        for item in items:
            self._doc.add_paragraph(item, style=style)
        return self

    def add_note(self, text: str, label: str = "Note") -> "DocBuilder":
        """Add an indented note paragraph."""
        p = self._doc.add_paragraph()
        r = p.add_run(f"{label}: ")
        r.bold = True
        r.font.color.rgb = _rgb(self._accent)
        p.add_run(text)
        p.paragraph_format.left_indent = Cm(1)
        return self

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------

    def add_table(
        self,
        headers: list[str],
        rows: list[list[str]],
        header_bg: str = "1F3864",
        header_fg: str = "FFFFFF",
        alt_row_bg: str | None = None,
    ) -> "DocBuilder":
        """
        Add a styled table.

        Parameters
        ----------
        headers:
            Column header strings.
        rows:
            List of row value lists.
        header_bg:
            Background hex for the header row (default dark-navy).
        header_fg:
            Foreground hex for header text (default white).
        alt_row_bg:
            If provided, alternating body rows get this background colour.
        """
        doc = self._doc
        tbl = doc.add_table(rows=1, cols=len(headers))
        tbl.style = self._TABLE_STYLE
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header
        hrow = tbl.rows[0]
        for i, h in enumerate(headers):
            cell = hrow.cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(h)
            run.bold = True
            run.font.color.rgb = _rgb(header_fg)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_cell_bg(cell, header_bg)

        # Body rows
        for row_idx, values in enumerate(rows):
            row = tbl.add_row()
            bg = alt_row_bg if (alt_row_bg and row_idx % 2 == 1) else None
            for col_idx, val in enumerate(values):
                cell = row.cells[col_idx]
                cell.text = str(val)
                if bg:
                    _set_cell_bg(cell, bg)

        doc.add_paragraph()
        return self

    # ------------------------------------------------------------------
    # Charts / diagrams
    # ------------------------------------------------------------------

    def add_bar_chart(
        self,
        title: str,
        categories: list[str],
        series: dict[str, list[float]],
        ylabel: str = "",
        width_inches: float = 6.0,
        caption: str = "",
    ) -> "DocBuilder":
        """
        Embed a grouped bar chart.

        Parameters
        ----------
        title:
            Chart title (shown above the chart).
        categories:
            Labels on the X-axis.
        series:
            Dict mapping series-name → list of values (one per category).
        ylabel:
            Optional Y-axis label.
        width_inches:
            Rendered width in the document.
        caption:
            Optional figure caption (italicised, centred).
        """
        n = len(categories)
        n_series = len(series)
        x = np.arange(n)
        width = 0.8 / max(n_series, 1)

        fig, ax = plt.subplots(figsize=(max(6, n * 1.2), 4))
        ax.set_facecolor("#F5F5F5")
        fig.patch.set_facecolor("#F5F5F5")

        colours = plt.rcParams["axes.prop_cycle"].by_key()["color"]
        for idx, (name, values) in enumerate(series.items()):
            offset = (idx - (n_series - 1) / 2) * width
            ax.bar(x + offset, values, width, label=name,
                   color=colours[idx % len(colours)], zorder=3)

        ax.set_xticks(x)
        ax.set_xticklabels(categories, fontsize=9)
        if ylabel:
            ax.set_ylabel(ylabel, fontsize=9)
        ax.set_title(title, fontsize=11, fontweight="bold",
                     color=f"#{self._accent}")
        if n_series > 1:
            ax.legend(fontsize=8.5)
        ax.grid(axis="y", linestyle="--", alpha=0.5, zorder=0)
        fig.tight_layout()

        self._embed_figure(fig, width_inches, caption)
        return self

    def add_line_chart(
        self,
        title: str,
        categories: list[str],
        series: dict[str, list[float]],
        ylabel: str = "",
        width_inches: float = 6.0,
        caption: str = "",
    ) -> "DocBuilder":
        """Embed a line chart."""
        fig, ax = plt.subplots(figsize=(max(6, len(categories) * 0.8), 4))
        ax.set_facecolor("#F5F5F5")
        fig.patch.set_facecolor("#F5F5F5")

        for name, values in series.items():
            ax.plot(categories, values, marker="o", linewidth=2, label=name)

        if ylabel:
            ax.set_ylabel(ylabel, fontsize=9)
        ax.set_title(title, fontsize=11, fontweight="bold",
                     color=f"#{self._accent}")
        ax.legend(fontsize=8.5)
        ax.grid(linestyle="--", alpha=0.5)
        fig.tight_layout()

        self._embed_figure(fig, width_inches, caption)
        return self

    def add_pie_chart(
        self,
        title: str,
        labels: list[str],
        values: list[float],
        width_inches: float = 5.0,
        caption: str = "",
    ) -> "DocBuilder":
        """Embed a pie chart."""
        fig, ax = plt.subplots(figsize=(5, 4))
        fig.patch.set_facecolor("#F5F5F5")
        ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
        ax.set_title(title, fontsize=11, fontweight="bold",
                     color=f"#{self._accent}")
        fig.tight_layout()

        self._embed_figure(fig, width_inches, caption)
        return self

    # ------------------------------------------------------------------
    # Image helpers
    # ------------------------------------------------------------------

    def _embed_figure(
        self,
        fig: plt.Figure,
        width_inches: float,
        caption: str,
    ) -> None:
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        buf.seek(0)
        plt.close(fig)
        self._doc.add_picture(buf, width=Inches(width_inches))
        self._doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            p = self._doc.add_paragraph(caption, style="Caption")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_image(
        self,
        path: str,
        width_inches: float = 5.0,
        caption: str = "",
    ) -> "DocBuilder":
        """Embed an existing image file."""
        self._doc.add_picture(path, width=Inches(width_inches))
        self._doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            p = self._doc.add_paragraph(caption, style="Caption")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return self

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save(self, output_path: str) -> None:
        """Write the document to *output_path*."""
        self._doc.save(output_path)
        print(f"Saved: {output_path}")


# ---------------------------------------------------------------------------
# Example document (run with --example)
# ---------------------------------------------------------------------------

def _build_example(output_path: str) -> None:
    b = DocBuilder(
        title="Example Report",
        subtitle="Generated with tools/generate_docx.py",
        author="FreshTake Team",
        date="March 2026",
    )

    b.add_section("1. Introduction")
    b.add_body(
        "This document was produced automatically by the generate_docx.py script. "
        "It demonstrates the main building blocks available: headings, paragraphs, "
        "bullet lists, tables, bar charts, line charts, and pie charts."
    )

    b.add_section("2. Bullet Lists & Notes")
    b.add_bullet_list([
        "python-docx handles the Word XML format",
        "matplotlib renders charts as embedded PNG images",
        "DocBuilder exposes a fluent (chained) API",
    ])
    b.add_note(
        "All helper functions are importable from tools.generate_docx "
        "so you can reuse them in your own scripts."
    )

    b.add_section("3. Tables")
    b.add_body("Below is a sample project-status table:")
    b.add_table(
        headers=["Task", "Owner", "Due Date", "Status"],
        rows=[
            ["Initial setup",  "Alice", "2026-04-01", "Done"],
            ["Write tests",    "Bob",   "2026-04-15", "In Progress"],
            ["Deploy to prod", "Carol", "2026-05-01", "Open"],
            ["Post-launch review", "Alice", "2026-05-15", "Open"],
        ],
        alt_row_bg="BDD7EE",
    )

    b.add_section("4. Bar Chart")
    b.add_bar_chart(
        title="Sprint Velocity",
        categories=["Sprint 1", "Sprint 2", "Sprint 3", "Sprint 4"],
        series={
            "Story Points Completed": [34, 41, 38, 46],
            "Story Points Planned":   [40, 40, 40, 45],
        },
        ylabel="Points",
        caption="Figure 1 – Sprint velocity over the last four sprints.",
    )

    b.add_section("5. Line Chart")
    b.add_line_chart(
        title="Bug Count Over Time",
        categories=["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"],
        series={
            "Open Bugs":   [12, 18, 15, 10, 6],
            "Closed Bugs": [0,   5,  9, 14, 18],
        },
        ylabel="Count",
        caption="Figure 2 – Bug trend across five weeks.",
    )

    b.add_section("6. Pie Chart")
    b.add_pie_chart(
        title="Work Distribution by Team",
        labels=["Backend", "Frontend", "QA", "DevOps"],
        values=[40, 30, 20, 10],
        caption="Figure 3 – Percentage of effort per team.",
    )

    b.add_section("7. Conclusion")
    b.add_body(
        "The generate_docx.py script can be used as a starting point for any new "
        "document. Simply import DocBuilder, chain together the helper methods, "
        "and call .save() with your desired output filename."
    )

    b.save(output_path)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate a .docx file using DocBuilder."
    )
    parser.add_argument(
        "--example",
        action="store_true",
        help="Build the built-in example document.",
    )
    parser.add_argument(
        "--output",
        default="example.docx",
        help="Output file path (default: example.docx).",
    )
    args = parser.parse_args()

    # --example is optional; running without it also produces the example document.
    _build_example(args.output)
