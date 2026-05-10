import json


def render(name, data, fmt="text"):
    fn = _RENDERERS.get(name)
    if not fn:
        return str(data)
    return fn(data, fmt)


def _cell(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    val = data["value"] if data["value"] else "(empty)"
    line = f"[{data['sheet']}] {data['ref']} = {val} ({data['type']})"
    parts = [line]
    if "table" in data and data["table"] is not None:
        t = data["table"]
        parts.append(f"  Table: {t['name']} ({t['range']})")
        if "column" in t:
            parts.append(f"  Column: \"{t['column']}\"")
        if "row" in t:
            parts.append(f"  Row: \"{t['row']}\"")
    if "merged_range" in data:
        parts.append(f"  Merged: {data['merged_range']}")
    if "style" in data:
        s = data["style"]
        f = s["font"]
        size = f"{f['size']} " if f["size"] else ""
        name = f"{f['name']} " if f["name"] else ""
        parts.append(f"  Font: {'bold ' if f['bold'] else ''}{'italic ' if f['italic'] else ''}{size}{name}#{f['color'] or 'default'}")
        fi = s["fill"]
        parts.append(f"  Fill: fg=#{fi['fg'] if fi['fg'] and fi['fg'] != '00000000' else 'none'} bg=#{fi['bg'] if fi['bg'] and fi['bg'] != '00000000' else 'none'}")
        b = s["border"]
        edges = [f"{k}={v}" for k, v in b.items() if v]
        parts.append(f"  Border: {' '.join(edges) or 'none'}")
        a = s["align"]
        parts.append(f"  Align: h={a['h'] or 'general'} v={a['v'] or 'bottom'} wrap={'yes' if a['wrap'] else 'no'}")
        parts.append(f"  Format: {s['number_format'] or 'General'}")
    return "\n".join(parts)


def _sheets(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    lines = []
    for s in data:
        tag = "(hidden)" if s["hidden"] else "visible"
        lines.append(f"{s['name']} ({tag}) — {s['rows']} rows, {s['cols']} cols")
    return "\n".join(lines) if lines else "(no sheets)"


def _info(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    return f"File: {data['filename']} | Size: {data['size']:,} bytes | Sheets: {data['sheets']}"


def _sheet(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    return f"Now on: {data['sheet']}"


def _help(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    lines = [f"{c['name']:10s} — {c['help']}" for c in data]
    return "\n".join(lines)


def _table(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    if isinstance(data, list):
        if not data:
            return "No tables registered"
        return "\n".join(f"{t['name']} ({t['range']}) [header={t.get('header', 'row')}]" for t in data)
    if isinstance(data, dict) and "message" in data:
        return data["message"]
    return str(data)


def _trace(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    lines = [f"[{data['sheet']}] {data['ref']} = {data['formula']} (formula)"]
    for d in data["dependencies"]:
        if d["type"] == "range":
            lines.append(f"  {d['ref']} (range)")
        else:
            v = d["value"] if d["value"] else "(empty)"
            lines.append(f"  {d['ref']} = {v} ({d['type']})")
    lines.append(f"  {data['ref']} = {data['computed']} (computed)")
    return "\n".join(lines)


def _find(data, fmt):
    if fmt == "json":
        return json.dumps(data, indent=2)
    if isinstance(data, list):
        if not data:
            return "No matches found"
        lines = []
        for d in data:
            v = d["value"] if d["value"] else "(empty)"
            line = f"  {d['ref']} = {v} ({d['type']})"
            if d.get("merged_range"):
                line += f" [Merged: {d['merged_range']}]"
            lines.append(line)
        return f"{len(data)} match(es)\n" + "\n".join(lines)
    if isinstance(data, dict) and "message" in data:
        return data["message"]
    return str(data)


_RENDERERS = {
    "cell": _cell,
    "sheets": _sheets,
    "info": _info,
    "sheet": _sheet,
    "help": _help,
    "table": _table,
    "trace": _trace,
    "find": _find,
}
