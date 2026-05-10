import json
import os
from datetime import datetime, date

from openpyxl import load_workbook, utils

_TYPE_LABELS = {
    "n": "number",
    "s": "text",
    "f": "formula",
    "d": "date",
    "b": "boolean",
    "e": "error",
    "nul": "empty",
}


class Session:
    def __init__(self, path):
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        self.path = path
        self.wb = load_workbook(path, data_only=False)
        self._current_sheet = self.wb.sheetnames[0]
        self._tables_path = path + ".tables.json"
        self._tables = {}
        self._load_tables()

    @property
    def current_sheet(self):
        return self._current_sheet

    def close(self):
        self.wb.close()

    def info(self):
        return {
            "filename": os.path.basename(self.path),
            "size": os.path.getsize(self.path),
            "sheets": len(self.wb.sheetnames),
        }

    def sheets(self):
        result = []
        for name in self.wb.sheetnames:
            ws = self.wb[name]
            first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
            result.append({
                "name": name,
                "hidden": ws.sheet_state == "hidden",
                "rows": ws.max_row or 0,
                "cols": len(first),
            })
        return result

    def switch_sheet(self, name):
        if name not in self.wb.sheetnames:
            raise ValueError(
                f"Sheet '{name}' not found. Available: {', '.join(self.wb.sheetnames)}"
            )
        self._current_sheet = name

    def cell(self, ref, sheet=None):
        cell = self._resolve_cell(ref, sheet)
        return self._fmt(cell.value), _TYPE_LABELS.get(cell.data_type, "unknown")

    def cell_computed(self, ref, sheet=None):
        sheet = sheet or self._current_sheet
        wb = load_workbook(self.path, data_only=True)
        try:
            ws = wb[sheet]
            cell = self._resolve_cell(ref, sheet, wb)
            return self._fmt(cell.value)
        finally:
            wb.close()

    def cell_style(self, ref, sheet=None):
        cell = self._resolve_cell(ref, sheet)
        f = cell.font
        fi = cell.fill
        b = cell.border
        a = cell.alignment
        return {
            "font": {
                "bold": f.bold,
                "italic": f.italic,
                "size": f.size,
                "name": f.name,
                "color": str(f.color.rgb) if f.color and f.color.rgb else None,
            },
            "fill": {
                "fg": str(fi.fgColor.rgb) if fi.fgColor and fi.fgColor.rgb else None,
                "bg": str(fi.bgColor.rgb) if fi.bgColor and fi.bgColor.rgb else None,
            },
            "border": {
                "top": b.top.style if b.top else None,
                "bottom": b.bottom.style if b.bottom else None,
                "left": b.left.style if b.left else None,
                "right": b.right.style if b.right else None,
            },
            "align": {
                "h": a.horizontal if a else None,
                "v": a.vertical if a else None,
                "wrap": a.wrap_text if a else None,
            },
            "number_format": cell.number_format,
        }

    def cell_table(self, ref, sheet=None):
        sheet = sheet or self._current_sheet
        ws = self.wb[sheet]
        col, row = self._parse_ref(ref)
        for tname, tref in ws.tables.items():
            bounds = utils.range_boundaries(tref)
            if bounds and bounds[0] <= col <= bounds[2] and bounds[1] <= row <= bounds[3]:
                header_cell = ws.cell(row=bounds[1], column=col)
                return {
                    "name": tname,
                    "range": tref,
                    "column": self._fmt(header_cell.value),
                }
        for tname, meta in self._tables.get(sheet, {}).items():
            if isinstance(meta, str):
                continue
            bounds = utils.range_boundaries(meta["range"])
            if bounds and bounds[0] <= col <= bounds[2] and bounds[1] <= row <= bounds[3]:
                result = {"name": tname, "range": meta["range"]}
                h = meta.get("header", "row")
                if h in ("row", "both"):
                    result["column"] = self._fmt(ws.cell(row=bounds[1], column=col).value)
                if h in ("column", "both"):
                    result["row"] = self._fmt(ws.cell(row=row, column=bounds[0]).value)
                return result
        return None

    def register_table(self, name, ref, sheet=None, header="row"):
        sheet = sheet or self._current_sheet
        try:
            utils.range_boundaries(ref)
        except (ValueError, AttributeError):
            raise ValueError(f"Invalid range: {ref}")
        self._tables.setdefault(sheet, {})[name] = {"range": ref.upper(), "header": header}
        self._save_tables()

    def unregister_table(self, name, sheet=None):
        sheet = sheet or self._current_sheet
        result = self._tables.get(sheet, {}).pop(name, None) is not None
        if result:
            self._save_tables()
        return result

    def list_tables(self, sheet=None):
        sheet = sheet or self._current_sheet
        result = []
        for name, meta in self._tables.get(sheet, {}).items():
            if isinstance(meta, str):
                continue
            result.append({"name": name, "range": meta["range"], "header": meta.get("header", "row"), "sheet": sheet})
        return result

    def clear_tables(self, sheet=None):
        sheet = sheet or self._current_sheet
        self._tables.pop(sheet, None)
        self._save_tables()

    def _load_tables(self):
        try:
            with open(self._tables_path) as f:
                raw = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self._tables = {}
            return
        self._tables = {}
        for sheet, tables in raw.items():
            if not isinstance(tables, dict):
                continue
            self._tables[sheet] = {}
            for name, meta in tables.items():
                if isinstance(meta, dict) and "range" in meta:
                    self._tables[sheet][name] = meta

    def _save_tables(self):
        with open(self._tables_path, "w") as f:
            json.dump(self._tables, f, indent=2)

    def _resolve_cell(self, ref, sheet, wb=None):
        wb = wb or self.wb
        sheet = sheet or self._current_sheet
        if sheet not in self.wb.sheetnames:
            raise ValueError(f"Sheet '{sheet}' not found.")
        ws = wb[sheet]
        col, row = self._parse_ref(ref)
        return ws.cell(row=row, column=col)

    @staticmethod
    def _parse_ref(ref):
        ref = ref.upper()
        col_str = "".join(c for c in ref if c.isalpha())
        row_str = "".join(c for c in ref if c.isdigit())
        if not col_str or not row_str:
            raise ValueError(f"Invalid cell reference: {ref}")
        return utils.column_index_from_string(col_str), int(row_str)

    @staticmethod
    def _fmt(v):
        if v is None:
            return ""
        if isinstance(v, (datetime, date)):
            return v.isoformat()
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
