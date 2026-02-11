"""
HTML report generator for IHSDM crash prediction summaries.
Adapted from hntb_report.py -- self-contained for PyInstaller builds.
"""
import base64
import os
from datetime import date


# HNTB Corporation logo (SVG, public domain) -- embedded so no file I/O needed
_DEFAULT_HNTB_LOGO_B64 = (
    "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9Im5vIj8+"
    "CjwhLS0gQ3JlYXRlZCB3aXRoIElua3NjYXBlIChodHRwOi8vd3d3Lmlua3NjYXBlLm9yZy8p"
    "IC0tPgoKPHN2ZwogICB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEu"
    "MS8iCiAgIHhtbG5zOmNjPSJodHRwOi8vY3JlYXRpdmVjb21tb25zLm9yZy9ucyMiCiAgIHht"
    "bG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyIK"
    "ICAgeG1sbnM6c3ZnPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIKICAgeG1sbnM9Imh0"
    "dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIgogICB2ZXJzaW9uPSIxLjAiCiAgIHdpZHRoPSIy"
    "NjAycHQiCiAgIGhlaWdodD0iNzU3cHQiCiAgIHZpZXdCb3g9IjAgMCAyNjAyIDc1NyIKICAg"
    "aWQ9InN2ZzIiPgogIDx0aXRsZQogICAgIGlkPSJ0aXRsZTI5ODgiPkhOVEIgbG9nbzwvdGl0"
    "bGU+CiAgPGRlZnMKICAgICBpZD0iZGVmczE4IiAvPgogIDxtZXRhZGF0YQogICAgIGlkPSJt"
    "ZXRhZGF0YTQiPgogICAgPHJkZjpSREY+CiAgICAgIDxjYzpXb3JrCiAgICAgICAgIHJkZjph"
    "Ym91dD0iIj4KICAgICAgICA8ZGM6Zm9ybWF0PmltYWdlL3N2Zyt4bWw8L2RjOmZvcm1hdD4K"
    "ICAgICAgICA8ZGM6dHlwZQogICAgICAgICAgIHJkZjpyZXNvdXJjZT0iaHR0cDovL3B1cmwu"
    "b3JnL2RjL2RjbWl0eXBlL1N0aWxsSW1hZ2UiIC8+CiAgICAgICAgPGRjOnRpdGxlPkhOVEIg"
    "bG9nbzwvZGM6dGl0bGU+CiAgICAgICAgPGNjOmxpY2Vuc2UKICAgICAgICAgICByZGY6cmVz"
    "b3VyY2U9Imh0dHA6Ly9jcmVhdGl2ZWNvbW1vbnMub3JnL2xpY2Vuc2VzL3B1YmxpY2RvbWFp"
    "bi8iIC8+CiAgICAgICAgPGRjOmNyZWF0b3I+CiAgICAgICAgICA8Y2M6QWdlbnQ+CiAgICAg"
    "ICAgICAgIDxkYzp0aXRsZT5Vc2VyOkpiYXJ0YSBhdCBlbi53aWtpcGVkaWEub3JnPC9kYzp0"
    "aXRsZT4KICAgICAgICAgIDwvY2M6QWdlbnQ+CiAgICAgICAgPC9kYzpjcmVhdG9yPgogICAg"
    "ICAgIDxkYzpyaWdodHM+CiAgICAgICAgICA8Y2M6QWdlbnQ+CiAgICAgICAgICAgIDxkYzp0"
    "aXRsZT5wdWJsaWMgZG9tYWluPC9kYzp0aXRsZT4KICAgICAgICAgIDwvY2M6QWdlbnQ+CiAg"
    "ICAgICAgPC9kYzpyaWdodHM+CiAgICAgICAgPGRjOmRlc2NyaXB0aW9uPmxvZ28gb2YgSE5U"
    "QiBDb3Jwb3JhdGlvbjwvZGM6ZGVzY3JpcHRpb24+CiAgICAgIDwvY2M6V29yaz4KICAgICAg"
    "PGNjOkxpY2Vuc2UKICAgICAgICAgcmRmOmFib3V0PSJodHRwOi8vY3JlYXRpdmVjb21tb25z"
    "Lm9yZy9saWNlbnNlcy9wdWJsaWNkb21haW4vIj4KICAgICAgICA8Y2M6cGVybWl0cwogICAg"
    "ICAgICAgIHJkZjpyZXNvdXJjZT0iaHR0cDovL2NyZWF0aXZlY29tbW9ucy5vcmcvbnMjUmVw"
    "cm9kdWN0aW9uIiAvPgogICAgICAgIDxjYzpwZXJtaXRzCiAgICAgICAgICAgcmRmOnJlc291"
    "cmNlPSJodHRwOi8vY3JlYXRpdmVjb21tb25zLm9yZy9ucyNEaXN0cmlidXRpb24iIC8+CiAg"
    "ICAgICAgPGNjOnBlcm1pdHMKICAgICAgICAgICByZGY6cmVzb3VyY2U9Imh0dHA6Ly9jcmVh"
    "dGl2ZWNvbW1vbnMub3JnL25zI0Rlcml2YXRpdmVXb3JrcyIgLz4KICAgICAgPC9jYzpMaWNl"
    "bnNlPgogICAgPC9yZGY6UkRGPgogIDwvbWV0YWRhdGE+CiAgPGcKICAgICB0cmFuc2Zvcm09"
    "Im1hdHJpeCgwLjEsMCwwLC0wLjEsMCw3NTcpIgogICAgIGlkPSJnNiIKICAgICBzdHlsZT0i"
    "ZmlsbDojMDAwMDAwO3N0cm9rZTpub25lIj4KICAgIDxwYXRoCiAgICAgICBkPSJtIDI4MCwy"
    "NDAgMjA4MCwwIDAsMjYwMCAyMDgwLDAgMCwtMjYwMCAyMDgwLDAgMCw3MDgwIC0yMDgwLDAg"
    "MCwtMjM2MCAtMjA4MCwwIDAsMjM2MCAtMjA4MCwwIDAsLTIzNjAgMCwtMjEyMCB6IgogICAg"
    "ICAgaWQ9InBhdGg4IgogICAgICAgc3R5bGU9ImZpbGw6IzEwMTAxMDtmaWxsLW9wYWNpdHk6"
    "MSIgLz4KICAgIDxwYXRoCiAgICAgICBkPSJtIDc1MjAsMjQwIDIwODAsMCAwLDMxMjAgMjE2"
    "MCwtMzEyMCAxNjQwLDAgMCw3MDgwIC0yMDQwLDAgMCwtMzA0MCAtMjEyMCwzMDQwIC0xNzIw"
    "LDAgeiIKICAgICAgIGlkPSJwYXRoMTAiCiAgICAgICBzdHlsZT0iZmlsbDojMTAxMDEwO2Zp"
    "bGwtb3BhY2l0eToxIiAvPgogICAgPHBhdGgKICAgICAgIGQ9Im0gMTQwMDAsNTI4MCAxNjAw"
    "LDAgMCwtNTA0MCAyMDgwLDAgMCw1MDQwIDE2MDAsMCAwLDIwNDAgLTUyODAsMCB6IgogICAg"
    "ICAgaWQ9InBhdGgxMiIKICAgICAgIHN0eWxlPSJmaWxsOiMxMDEwMTA7ZmlsbC1vcGFjaXR5"
    "OjEiIC8+CiAgICA8cGF0aAogICAgICAgZD0ibSAxOTg4MCwyNDAgMTc2OCwwIGMgMTcxOCww"
    "IDE3ODIsMSAxOTM3LDIwIDE5NywyNSAzODIsNjEgNTM2LDEwNSA4NDcsMjQ0IDE0MTMsODIz"
    "IDE1OTQsMTYzMiA1MywyMzMgNTgsMjkyIDU5LDU3OCAwLDIzOSAtMiwyODEgLTIxLDM2NiAt"
    "NjEsMjc0IC0xNjUsNTA1IC0zMjIsNzE2IC04MSwxMDkgLTI0MCwyNzAgLTM0MywzNDcgLTQz"
    "LDMyIC03OCw2MSAtNzgsNjQgMCwzIDM5LDQ2IDg4LDk2IDE0NCwxNTAgMjI0LDI2NSAzMTMs"
    "NDUxIDU4LDEyMiAxMDAsMjQ2IDEzMSwzODUgMTksODkgMjIsMTM0IDIyLDI5NSAwLDI4OCAt"
    "NDUsNTM0IC0xNDAsNzc1IC0yODEsNzA1IC05NDQsMTE3MCAtMTg2MSwxMjQ1IC0zNjMsNSAt"
    "NzYzLDUgLTc2Myw1IGwgLTI5MjAsMCB6IG0gMzQwNSw1NDg2IGMgMjMxLC02OSAzNTgsLTIy"
    "MCAzODAsLTQ1MiAxOSwtMjAzIC05OSwtNDE4IC0yNzcsLTUwMyAtMTI1LC01OSAtMTY5LC01"
    "MSAtOTA4LC01MSBsIC02NDAsMCAwLDEwNDAgNjgwLDAgYyA2MzMsLTMgNzEwLC0xOCA3NjUs"
    "LTM0IHogbSAzNSwtMjYyNCBjIDEzNiwtMzYgMjI0LC04NCAzMDYsLTE2NiAxMTUsLTExNiAx"
    "NjQsLTI1MCAxNjQsLTQ1MSAwLC0zMDkgLTE0NSwtNTIxIC00MjIsLTYxNiAtNDYsLTE1IC0x"
    "MjUsLTI0IC0xNjgsLTI5IC00OSwtNyAtMzE3LDAgLTcyMCwwIGwgLTY0MCwwIDAsMTI4MCA2"
    "ODAsMCBjIDY5MCwtMyA3MTYsNCA4MDAsLTE4IHoiCiAgICAgICBpZD0icGF0aDE0IgogICAg"
    "ICAgc3R5bGU9ImZpbGw6IzEwMTAxMDtmaWxsLW9wYWNpdHk6MSIgLz4KICA8L2c+CiAgPHJl"
    "Y3QKICAgICB3aWR0aD0iMjA3Ljk5OTk0IgogICAgIGhlaWdodD0iMjExLjk5OTk4IgogICAg"
    "IHJ5PSIwIgogICAgIHg9IjI4IgogICAgIHk9IjI2MSIKICAgICBpZD0icmVjdDI5OTUiCiAg"
    "ICAgc3R5bGU9ImZpbGw6I2Y3NWEyMTtmaWxsLW9wYWNpdHk6MTtzdHJva2U6bm9uZSIgLz4K"
    "PC9zdmc+Cg=="
)


def _guess_mime(path_or_bytes):
    """Return MIME type string for logo embedding."""
    if isinstance(path_or_bytes, str):
        ext = os.path.splitext(path_or_bytes)[1].lower()
    else:
        ext = ""
    if ext == ".svg":
        return "image/svg+xml"
    if ext in (".jpg", ".jpeg"):
        return "image/jpeg"
    return "image/png"


def _embed_logo(logo_path):
    if logo_path and os.path.exists(logo_path):
        mime = _guess_mime(logo_path)
        with open(logo_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        return b64, mime
    return "", "image/png"


def _fmt(val, decimals=2):
    if val is None:
        return "\u2014"
    if abs(val) < 0.005 and decimals == 2:
        return "0.00"
    return f"{val:,.{decimals}f}"


def _fmt_diff(val, decimals=2):
    if val is None:
        return "\u2014", ""
    cls = "neg" if val < -0.005 else ("pos" if val > 0.005 else "zero")
    sign = "+" if val > 0.005 else ""
    return f"{sign}{val:,.{decimals}f}", cls


# ── Report class ─────────────────────────────────────────────────────────────


class Report:
    def __init__(
        self,
        title,
        subtitle="",
        project_id="",
        logo_path=None,
        logo_bytes=None,
        footer_text="",
    ):
        self.title = title
        self.subtitle = subtitle
        self.project_id = project_id
        self.logo_path = logo_path
        self.logo_bytes = logo_bytes
        self.footer_text = footer_text
        self.sections = []

    # ── Public API ────────────────────────────────────────────────────────

    def add_note(self, text):
        self.sections.append(("note", {"text": text}))

    def add_bar_chart(self, bars):
        self.sections.append(("bar_chart", {"bars": bars}))

    def add_table(self, title, columns, rows, total_row=None, distribution_fn=None):
        self.sections.append(("table", {
            "title": title,
            "columns": columns,
            "rows": rows,
            "total_row": total_row,
            "distribution_fn": distribution_fn,
        }))

    def add_diff_table(self, title, columns, rows, total_row=None, bar_key=None):
        self.sections.append(("diff_table", {
            "title": title,
            "columns": columns,
            "rows": rows,
            "total_row": total_row,
            "bar_key": bar_key,
        }))

    def add_side_by_side_tables(self, tables):
        self.sections.append(("side_by_side", {"tables": tables}))

    def add_separator(self):
        self.sections.append(("separator", {}))

    def add_metric_cards(self, cards):
        self.sections.append(("metric_cards", {"cards": cards}))

    # ── Generate ──────────────────────────────────────────────────────────

    def to_html(self):
        logo_b64 = ""
        logo_mime = "image/svg+xml"

        if self.logo_bytes:
            logo_b64 = base64.b64encode(self.logo_bytes).decode("utf-8")
            logo_mime = "image/png"
        elif self.logo_path:
            logo_b64, logo_mime = _embed_logo(self.logo_path)

        # Fall back to the embedded HNTB logo
        if not logo_b64:
            logo_b64 = _DEFAULT_HNTB_LOGO_B64
            logo_mime = "image/svg+xml"

        today_str = date.today().strftime("%B %d, %Y")

        if logo_b64:
            logo_tag = f'<img src="data:{logo_mime};base64,{logo_b64}" alt="Logo" />'
        else:
            logo_tag = '<span style="font-weight:800;">HNTB</span>'

        body_html = "\n".join(self._render_section(s) for s in self.sections)

        footer_line = self.footer_text or self.subtitle

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{self.title}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet" />
<style>
{self._css()}
</style>
</head>
<body>

<button class="print-btn no-print" onclick="window.print()">Print / Save PDF</button>

<div class="page">
  <div class="header">
    <div>
      <h1>{self.title}</h1>
      <div class="subtitle">{self.subtitle}</div>
    </div>
    <div class="meta">
      {logo_tag}<br>
      <div class="date-info">
        Generated: {today_str}<br>
        {self.project_id}
      </div>
    </div>
  </div>

  <div class="body">
    {body_html}
  </div>

  <div class="footer">
    <div class="footer-left">
      {logo_tag}
    </div>
    <div class="footer-text">
      {footer_line}<br>
      {self.project_id}
    </div>
  </div>
</div>

</body>
</html>"""

    def generate(self, output_path):
        html = self.to_html()
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

    # ── Section renderers ─────────────────────────────────────────────────

    def _render_section(self, section):
        stype, data = section
        if stype == "note":
            return f'<div class="note">{data["text"]}</div>'
        if stype == "bar_chart":
            return self._render_bar_chart(data)
        if stype == "table":
            return self._render_table(data)
        if stype == "diff_table":
            return self._render_diff_table(data)
        if stype == "side_by_side":
            return self._render_side_by_side(data)
        if stype == "separator":
            return '<hr class="section-break" />'
        if stype == "metric_cards":
            return self._render_metric_cards(data)
        return ""

    def _render_bar_chart(self, data):
        bars = data["bars"]
        max_val = max((b["value"] for b in bars), default=0)
        if max_val == 0:
            return ""
        rows_html = ""
        for b in bars:
            bar_pct = (b["value"] / max_val) * 100
            segments = b.get("segments", [])
            if segments:
                total = sum(s[0] for s in segments)
                segs_html = ""
                for seg_val, seg_cls in segments:
                    seg_pct = (seg_val / total) * bar_pct if total else 0
                    if seg_pct > 0.1:
                        segs_html += f'<div class="bar-seg {seg_cls}" style="width:{seg_pct:.2f}%"></div>'
            else:
                segs_html = f'<div class="bar-seg default" style="width:{bar_pct:.1f}%"></div>'
            decimals = b.get("decimals", 1)
            rows_html += (
                f'<div class="bar-row">'
                f'<span class="bar-label">{b["label"]}</span>'
                f'<div class="bar-track">{segs_html}</div>'
                f'<span class="bar-value">{_fmt(b["value"], decimals)}</span>'
                f"</div>"
            )
        return f'<div class="bar-chart">{rows_html}</div>'

    def _render_table(self, data):
        title = data["title"]
        columns = data["columns"]
        rows = data["rows"]
        total_row = data.get("total_row")
        dist_fn = data.get("distribution_fn")

        header_cells = "".join(
            f'<th class="{"cat-col" if i == 0 else ""}">{c["header"]}</th>'
            for i, c in enumerate(columns)
        )
        if dist_fn:
            header_cells += '<th class="bar-col">Distribution</th>'

        body = ""
        for r in rows:
            cells = ""
            for i, c in enumerate(columns):
                val = r.get(c["key"])
                decimals = c.get("decimals", 2)
                if i == 0:
                    cells += f'<td class="cat-col">{val if val is not None else ""}</td>'
                else:
                    cells += f"<td>{_fmt(val, decimals)}</td>"
            if dist_fn:
                cells += f"<td>{dist_fn(r)}</td>"
            body += f"<tr>{cells}</tr>"

        if total_row:
            cells = ""
            for i, c in enumerate(columns):
                val = total_row.get(c["key"])
                decimals = c.get("decimals", 2)
                if i == 0:
                    cells += f'<td class="cat-col">{val if val is not None else "TOTAL"}</td>'
                else:
                    cells += f"<td>{_fmt(val, decimals)}</td>"
            if dist_fn:
                cells += f"<td>{dist_fn(total_row)}</td>"
            body += f'<tr class="total-row">{cells}</tr>'

        return (
            f'<div class="table-section">'
            f'<div class="table-title">{title}</div>'
            f"<table><thead><tr>{header_cells}</tr></thead>"
            f"<tbody>{body}</tbody></table></div>"
        )

    def _render_diff_table(self, data):
        title = data["title"]
        columns = data["columns"]
        rows = data["rows"]
        total_row = data.get("total_row")
        bar_key = data.get("bar_key")

        max_abs = 0
        if bar_key:
            all_vals = [abs(r.get(bar_key, 0) or 0) for r in rows]
            if total_row:
                all_vals.append(abs(total_row.get(bar_key, 0) or 0))
            max_abs = max(all_vals) if all_vals else 0

        header_cells = "".join(
            f'<th class="{"cat-col" if i == 0 else ""}">{c["header"]}</th>'
            for i, c in enumerate(columns)
        )
        if bar_key:
            header_cells += '<th class="bar-col">Impact</th>'

        body = ""
        for r in rows:
            cells = ""
            for i, c in enumerate(columns):
                val = r.get(c["key"])
                decimals = c.get("decimals", 2)
                if i == 0:
                    cells += f'<td class="cat-col">{val if val is not None else ""}</td>'
                else:
                    v, cls = _fmt_diff(val, decimals)
                    cells += f'<td class="{cls}">{v}</td>'
            if bar_key:
                cells += f"<td>{self._diff_bar(r.get(bar_key, 0), max_abs)}</td>"
            body += f"<tr>{cells}</tr>"

        if total_row:
            cells = ""
            for i, c in enumerate(columns):
                val = total_row.get(c["key"])
                decimals = c.get("decimals", 2)
                if i == 0:
                    cells += f'<td class="cat-col">{val if val is not None else "TOTAL"}</td>'
                else:
                    v, cls = _fmt_diff(val, decimals)
                    cells += f'<td class="{cls}">{v}</td>'
            if bar_key:
                cells += f"<td>{self._diff_bar(total_row.get(bar_key, 0), max_abs)}</td>"
            body += f'<tr class="total-row">{cells}</tr>'

        return (
            f'<div class="table-section">'
            f'<div class="table-title">{title}</div>'
            f"<table><thead><tr>{header_cells}</tr></thead>"
            f"<tbody>{body}</tbody></table></div>"
        )

    def _render_side_by_side(self, data):
        tables = data["tables"]
        n = len(tables)
        cols_css = " ".join(["1fr"] * n)
        tables_html = ""
        for t in tables:
            columns = t["columns"]
            rows = t["rows"]
            total_row = t.get("total_row")
            header_cells = "".join(f'<th>{c["header"]}</th>' for c in columns)
            body = ""
            for r in rows:
                is_blank = r.get("blank", False)
                row_cls = ' class="blank-row"' if is_blank else ""
                cells = ""
                for i, c in enumerate(columns):
                    val = r.get(c["key"])
                    decimals = c.get("decimals", 2)
                    if is_blank and i > 0:
                        cells += "<td>&mdash;</td>"
                    elif i == 0:
                        cells += f'<td>{val if val is not None else ""}</td>'
                    else:
                        cells += f"<td>{_fmt(val, decimals)}</td>"
                body += f"<tr{row_cls}>{cells}</tr>"
            if total_row:
                cells = ""
                for i, c in enumerate(columns):
                    val = total_row.get(c["key"])
                    decimals = c.get("decimals", 2)
                    if i == 0:
                        cells += f'<td>{val if val is not None else "Total"}</td>'
                    else:
                        cells += f"<td>{_fmt(val, decimals)}</td>"
                body += f'<tr class="total-row">{cells}</tr>'
            tables_html += (
                f'<div class="int-table-section">'
                f'<div class="int-table-title">{t["title"]}</div>'
                f'<table class="int-table"><thead><tr>{header_cells}</tr></thead>'
                f"<tbody>{body}</tbody></table></div>"
            )
        return f'<div class="side-by-side" style="grid-template-columns:{cols_css}">{tables_html}</div>'

    def _render_metric_cards(self, data):
        cards_html = ""
        for c in data["cards"]:
            style = c.get("style", "default")
            pct_html = ""
            if "pct" in c and c["pct"]:
                direction = c.get("pct_direction", "flat")
                pct_html = f'<span class="pct {direction}">{c["pct"]}</span>'
            cards_html += (
                f'<div class="badge {style}">'
                f'<span class="label">{c["label"]}</span>'
                f'<span class="value">{c["value"]}</span>'
                f"{pct_html}</div>"
            )
        return f'<div class="badges">{cards_html}</div>'

    @staticmethod
    def _diff_bar(val, max_abs):
        empty = (
            '<div class="diff-bar">'
            '<div class="diff-half left"></div>'
            '<div class="diff-center"></div>'
            '<div class="diff-half right"></div></div>'
        )
        if max_abs == 0 or val is None:
            return empty
        pct = (abs(val) / max_abs) * 90
        if val < -0.005:
            return (
                '<div class="diff-bar">'
                f'<div class="diff-half left"><div class="diff-neg" style="width:{pct:.1f}%"></div></div>'
                '<div class="diff-center"></div>'
                '<div class="diff-half right"></div></div>'
            )
        if val > 0.005:
            return (
                '<div class="diff-bar">'
                '<div class="diff-half left"></div>'
                '<div class="diff-center"></div>'
                f'<div class="diff-half right"><div class="diff-pos" style="width:{pct:.1f}%"></div></div></div>'
            )
        return empty

    # ── CSS ───────────────────────────────────────────────────────────────

    @staticmethod
    def _css():
        return """
  :root {
    --brand-primary: #0e2141;
    --brand-accent: #f6a800;
    --ink: #1e2a37;
    --ink-muted: #6b7a90;
    --border: #e6ebf2;
    --sev-k: #fd281b;
    --sev-a: #fcad32;
    --sev-b: #fbff47;
    --sev-c: #316bf9;
    --sev-pd: #4dfe41;
  }
  @page { size: 8.5in 11in; margin: 0.2in; }
  * { box-sizing: border-box; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
  body { font-family: Inter, system-ui, sans-serif; color: var(--ink); margin: 0; padding: 0; background: #e8ecf2; }
  @media screen {
    .page { width: 8.5in; min-height: 11in; margin: 0.4in auto; background: #fff; box-shadow: 0 4px 24px rgba(0,0,0,.15); display: flex; flex-direction: column; }
    .no-print { display: block; }
  }
  @media print {
    html, body { background: #fff; margin: 0; padding: 0; }
    .page { width: 100%; margin: 0; box-shadow: none; display: flex; flex-direction: column; min-height: 100vh; }
    .no-print { display: none !important; }
  }
  .header { background: linear-gradient(135deg, var(--brand-primary) 0%, #1a3a5c 100%); color: #fff; padding: 0.15in 0.25in; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; }
  .header h1 { margin: 0; font-size: 1.15rem; font-weight: 800; letter-spacing: -0.02em; }
  .header .subtitle { font-size: 0.7rem; opacity: .9; margin-top: 2px; font-weight: 500; }
  .header .meta { text-align: right; }
  .header .meta img { height: 26px; margin-bottom: 2px; }
  .header .date-info { font-size: 0.58rem; opacity: .85; line-height: 1.4; }
  .body { flex: 1; padding: 0.08in 0.2in 0.06in; display: flex; flex-direction: column; gap: 0.04in; }
  .note { font-size: 0.6rem; font-weight: 600; color: var(--ink-muted); font-style: italic; }
  .badges { display: flex; gap: 10px; align-items: center; margin-bottom: 0px; }
  .badge { display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 4px; font-size: 0.62rem; font-weight: 700; }
  .badge .label { color: var(--ink-muted); }
  .badge .value { font-size: 0.85rem; font-weight: 800; }
  .badge.default { background: #f0f4fa; border-left: 3px solid var(--brand-primary); }
  .badge.default .value { color: var(--brand-primary); }
  .badge.green { background: #f0faf4; border-left: 3px solid #16a34a; }
  .badge.green .value { color: #16a34a; }
  .badge.amber { background: #fef7ed; border-left: 3px solid #b45309; }
  .badge.amber .value { color: #b45309; }
  .badge.blue { background: #f0f4fa; border-left: 3px solid var(--brand-primary); }
  .badge.blue .value { color: var(--brand-primary); }
  .badge .pct { font-size: 0.6rem; font-weight: 700; padding: 1px 5px; border-radius: 3px; }
  .badge .pct.up { background: #fef2f2; color: #dc2626; }
  .badge .pct.down { background: #f0fdf4; color: #16a34a; }
  .badge .pct.flat { background: #f8fafc; color: var(--ink-muted); }
  .bar-chart { display: flex; flex-direction: column; gap: 4px; padding: 4px 0; }
  .bar-row { display: flex; align-items: center; gap: 6px; }
  .bar-label { font-size: 0.58rem; font-weight: 700; width: 70px; text-align: right; flex-shrink: 0; color: var(--ink); }
  .bar-track { flex: 1; height: 18px; background: #f0f3f7; border-radius: 3px; overflow: hidden; position: relative; }
  .bar-seg { height: 100%; display: inline-block; float: left; }
  .bar-seg.k { background: var(--sev-k); }
  .bar-seg.a { background: var(--sev-a); }
  .bar-seg.b { background: var(--sev-b); }
  .bar-seg.c { background: var(--sev-c); }
  .bar-seg.pd { background: var(--sev-pd); }
  .bar-seg.default { background: linear-gradient(90deg, var(--brand-primary), #1a3a5c); }
  .bar-seg.green { background: linear-gradient(90deg, #16a34a, #22c55e); }
  .bar-seg.amber { background: linear-gradient(90deg, #b45309, #d97706); }
  .bar-seg.blue { background: linear-gradient(90deg, #0e2141, #1a3a5c); }
  .bar-seg.c0 { background: #0e2141; }
  .bar-seg.c1 { background: #2563eb; }
  .bar-seg.c2 { background: #16a34a; }
  .bar-seg.c3 { background: #f59e0b; }
  .bar-seg.c4 { background: #dc2626; }
  .bar-seg.c5 { background: #8b5cf6; }
  .bar-seg.c6 { background: #06b6d4; }
  .bar-seg.c7 { background: #ec4899; }
  .bar-seg.c8 { background: #84cc16; }
  .bar-seg.c9 { background: #f97316; }
  .bar-seg:first-child { border-radius: 3px 0 0 3px; }
  .bar-seg:last-child { border-radius: 0 3px 3px 0; }
  .bar-value { font-size: 0.55rem; font-weight: 800; color: var(--ink); white-space: nowrap; margin-left: 6px; line-height: 18px; }
  .table-section { margin-bottom: 1px; }
  .table-title { font-size: 0.68rem; font-weight: 800; color: var(--brand-primary); text-transform: uppercase; letter-spacing: 0.04em; padding: 2px 0 1px; border-bottom: 2px solid var(--brand-primary); margin-bottom: 0; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 0.6rem; }
  th { background: #f8fafc; padding: 2px 5px; text-align: right; font-weight: 700; color: var(--ink-muted); border-bottom: 1.5px solid var(--border); font-size: 0.55rem; text-transform: uppercase; }
  th.cat-col { text-align: left; width: 35%; }
  th.bar-col { text-align: center; width: 20%; }
  td { padding: 2px 5px; text-align: right; border-bottom: 1px solid #f0f3f7; font-variant-numeric: tabular-nums; }
  td.cat-col { text-align: left; font-weight: 600; color: var(--ink); font-size: 0.58rem; }
  .total-row td { font-weight: 800; border-top: 1.5px solid var(--brand-primary); border-bottom: 2px solid var(--brand-primary); color: var(--brand-primary); padding: 3px 5px; }
  td.pos { color: #dc2626; }
  td.neg { color: #16a34a; }
  td.zero { color: var(--ink-muted); }
  .sev-bar { display: flex; height: 12px; border-radius: 3px; overflow: hidden; }
  .sev-bar .seg { min-width: 2px; }
  .sev-bar .seg.k { background: var(--sev-k); }
  .sev-bar .seg.a { background: var(--sev-a); }
  .sev-bar .seg.b { background: var(--sev-b); }
  .sev-bar .seg.c { background: var(--sev-c); }
  .sev-bar .seg.pd { background: var(--sev-pd); }
  .sev-bar .seg.c0 { background: #0e2141; }
  .sev-bar .seg.c1 { background: #2563eb; }
  .sev-bar .seg.c2 { background: #16a34a; }
  .sev-bar .seg.c3 { background: #f59e0b; }
  .sev-bar .seg.c4 { background: #dc2626; }
  .sev-bar .seg.c5 { background: #8b5cf6; }
  .section-break { border: none; border-top: 2px solid var(--brand-primary); margin: 0.06in 0; opacity: 0.3; }
  .diff-bar { display: flex; align-items: center; height: 12px; width: 100%; }
  .diff-half { width: 50%; height: 10px; display: flex; }
  .diff-half.left { justify-content: flex-end; }
  .diff-half.right { justify-content: flex-start; }
  .diff-center { width: 1px; height: 12px; background: var(--ink-muted); opacity: 0.5; flex-shrink: 0; }
  .diff-neg { height: 10px; background: #16a34a; border-radius: 3px 0 0 3px; }
  .diff-pos { height: 10px; background: #dc2626; border-radius: 0 3px 3px 0; }
  .side-by-side { display: grid; gap: 0.08in; }
  .int-table-title { font-size: 0.6rem; font-weight: 800; color: var(--brand-primary); text-transform: uppercase; letter-spacing: 0.03em; padding: 3px 0 1px; border-bottom: 2px solid var(--brand-primary); }
  .int-table { width: 100%; border-collapse: collapse; font-size: 0.55rem; }
  .int-table th { font-size: 0.5rem; padding: 2px 3px; }
  .int-table td { padding: 1.5px 3px; font-size: 0.53rem; }
  .int-table .total-row td { font-size: 0.55rem; }
  .int-table .blank-row td { color: var(--ink-muted); opacity: 0.4; }
  .footer { border-top: 1.5px solid var(--border); padding: 4px 0.25in; display: flex; align-items: center; justify-content: space-between; background: linear-gradient(135deg, #fbfdff 0%, #f8fafc 100%); flex-shrink: 0; }
  .footer-left { display: flex; align-items: center; gap: 8px; }
  .footer-left img { height: 20px; }
  .footer-text { font-size: 0.52rem; color: var(--ink-muted); text-align: right; line-height: 1.5; }
  .print-btn { position: fixed; bottom: 1rem; right: 1rem; background: var(--brand-primary); color: #fff; border: none; padding: .6rem 1.2rem; border-radius: 2rem; font-weight: 700; font-size: .85rem; cursor: pointer; box-shadow: 0 4px 16px rgba(0,0,0,.25); z-index: 1000; }
  .print-btn:hover { background: #163461; }
"""


# ── Convenience helpers ──────────────────────────────────────────────────────


def severity_bar(row, keys=None):
    if keys is None:
        keys = [("K", "k"), ("A", "a"), ("B", "b"), ("C", "c"), ("PD", "pd")]
    total = sum(row.get(k, 0) or 0 for k, _ in keys)
    if total == 0:
        return '<div class="sev-bar"></div>'
    parts = []
    for key, cls in keys:
        val = row.get(key, 0) or 0
        pct = (val / total) * 100
        if pct > 0.5:
            parts.append(f'<div class="seg {cls}" style="width:{max(pct, 3):.1f}%"></div>')
    return f'<div class="sev-bar">{"".join(parts)}</div>'


def generic_bar(row, columns, colors=None):
    if colors is None:
        colors = ["c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9"]
    total = sum(abs(row.get(c, 0) or 0) for c in columns)
    if total == 0:
        return '<div class="sev-bar"></div>'
    parts = []
    for i, col in enumerate(columns):
        val = abs(row.get(col, 0) or 0)
        pct = (val / total) * 100
        cls = colors[i % len(colors)]
        if pct > 0.5:
            parts.append(f'<div class="seg {cls}" style="width:{max(pct, 3):.1f}%"></div>')
    return f'<div class="sev-bar">{"".join(parts)}</div>'
