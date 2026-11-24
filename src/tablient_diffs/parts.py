"""Compare the workbook parts that have no cell text: charts, pivot tables, macros and images.

Charts and pivot tables are read far enough to say what changed (type, title, series ranges, pivot source and fields).
Macros and images are only detected as added, removed or changed. Anything that cannot be parsed falls back to a
one-line "changed" so a reviewer still knows to open the file.
"""
from __future__ import annotations

import hashlib
import io
import posixpath
import xml.etree.ElementTree as ET
import zipfile

KINDS = [
    ("Chart", "xl/charts/chart"),
    ("Pivot table", "xl/pivotTables/pivotTable"),
    ("Macros", "xl/vbaProject.bin"),
    ("Image", "xl/media/"),
]
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _hashes(data: bytes | None) -> dict[str, str]:
    if data is None:
        return {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: hashlib.sha1(z.read(n)).hexdigest() for n in z.namelist() if any(n.startswith(p) for _, p in KINDS)}


def _kind(name: str) -> str:
    return next(k for k, p in KINDS if name.startswith(p))


def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter(A + "t")).strip() if el is not None else ""


def _ref(el) -> str:
    """Cell range or literal behind a chart series element (name, categories or values)."""
    if el is None:
        return ""
    f = el.find(f".//{C}f")
    if f is not None and f.text:
        return f.text
    v = el.find(f".//{C}v")
    return (v.text or "") if v is not None else ""


def describe_chart(xml: bytes) -> dict:
    root = ET.fromstring(xml)
    plot = root.find(f".//{C}plotArea")
    if plot is None:
        raise ValueError("not a chart part")
    kinds, series = [], []
    for el in plot:
        if el.tag.endswith("Chart"):
            kind = el.tag[len(C):-len("Chart")]
            bar_dir = el.find(f"{C}barDir")
            if kind == "bar" and bar_dir is not None:
                kind = "column" if bar_dir.get("val") == "col" else "bar"
            kinds.append(kind)
            for ser in el.findall(f"{C}ser"):
                series.append({
                    "name": _ref(ser.find(f"{C}tx")),
                    "cat": _ref(ser.find(f"{C}cat")) or _ref(ser.find(f"{C}xVal")),
                    "val": _ref(ser.find(f"{C}val")) or _ref(ser.find(f"{C}yVal")),
                })
    return {"type": "+".join(kinds) or "unknown", "title": _text(root.find(f".//{C}chart/{C}title")), "series": series}


def _chart_label(d: dict) -> str:
    title = f" '{d['title']}'" if d["title"] else ""
    return f"{d['type']} chart{title} with {len(d['series'])} series"


def diff_chart(old: dict, new: dict) -> list[str]:
    out = []
    if old["type"] != new["type"]:
        out.append(f"type {old['type']} -> {new['type']}")
    if old["title"] != new["title"]:
        out.append(f"title '{old['title']}' -> '{new['title']}'")
    oldn, newn = len(old["series"]), len(new["series"])
    for i in range(max(oldn, newn)):
        o = old["series"][i] if i < oldn else None
        n = new["series"][i] if i < newn else None
        label = (n or o)["name"] or f"series {i + 1}"
        if o is None:
            out.append(f"series added: {label} ({n['val']})")
        elif n is None:
            out.append(f"series removed: {label} ({o['val']})")
        else:
            for field, word in (("name", "name"), ("cat", "categories"), ("val", "values")):
                if o[field] != n[field]:
                    out.append(f"series {label}: {word} {o[field] or '-'} -> {n[field] or '-'}")
    return out


def describe_pivot(z: zipfile.ZipFile, name: str) -> dict:
    root = ET.fromstring(z.read(name))
    if root.tag != S + "pivotTableDefinition":
        raise ValueError("not a pivot table part")
    names, source = [], ""
    rels = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
    if rels in z.namelist():
        for rel in ET.fromstring(z.read(rels)).findall(f"{R}Relationship"):
            target = posixpath.normpath(posixpath.join(posixpath.dirname(name), rel.get("Target", "")))
            if "pivotCacheDefinition" in target and target in z.namelist():
                cache = ET.fromstring(z.read(target))
                names = [f.get("name", "") for f in cache.iter(S + "cacheField")]
                ws = cache.find(f".//{S}worksheetSource")
                if ws is not None:
                    source = f"{ws.get('sheet') or ws.get('name') or ''}!{ws.get('ref', '')}".strip("!")

    def field(idx) -> str:
        i = int(idx)
        return names[i] if 0 <= i < len(names) else f"field {i}"

    def listed(tag, attr):
        el = root.find(S + tag)
        return [field(f.get(attr)) for f in el if f.get(attr) not in (None, "-2")] if el is not None else []

    data = []
    df = root.find(S + "dataFields")
    for d in (df if df is not None else []):
        data.append(d.get("name") or f"{d.get('subtotal', 'sum')} of {field(d.get('fld'))}")
    loc = root.find(S + "location")
    return {"name": root.get("name", ""), "location": loc.get("ref", "") if loc is not None else "", "source": source,
            "rows": listed("rowFields", "x"), "cols": listed("colFields", "x"), "pages": listed("pageFields", "fld"),
            "data": data}


def _pivot_label(d: dict) -> str:
    return f"pivot table '{d['name']}'" + (f" from {d['source']}" if d["source"] else "")


def diff_pivot(old: dict, new: dict) -> list[str]:
    out = []
    for key, word in (("source", "source"), ("location", "location")):
        if old[key] != new[key]:
            out.append(f"{word} {old[key] or '-'} -> {new[key] or '-'}")
    for key, word in (("rows", "rows"), ("cols", "columns"), ("pages", "filters"), ("data", "values")):
        if old[key] != new[key]:
            out.append(f"{word} {', '.join(old[key]) or '-'} -> {', '.join(new[key]) or '-'}")
    return out


def _described(data: bytes | None) -> dict[str, tuple[str, object]]:
    """{part name: (kind, description)} for the charts and pivot tables that could be parsed."""
    out: dict[str, tuple[str, object]] = {}
    if data is None:
        return out
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for n in z.namelist():
            try:
                if n.startswith("xl/charts/chart") and n.endswith(".xml"):
                    out[n] = ("Chart", describe_chart(z.read(n)))
                elif n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml"):
                    out[n] = ("Pivot table", describe_pivot(z, n))
            except Exception:  # an unusual part falls back to the generic line below
                continue
    return out


def diff_parts(old: bytes | None, new: bytes | None) -> list[str]:
    """Lines for the parts that were added, removed or changed. Parts are matched by name, so renumbering after a
    delete can show as a change."""
    a, b = _hashes(old), _hashes(new)
    da, db = _described(old), _described(new)
    lines: list[str] = []
    generic: dict[tuple[str, str], int] = {}

    def bump(kind, verb):
        generic[(kind, verb)] = generic.get((kind, verb), 0) + 1

    for n in list(b) + [n for n in a if n not in b]:
        kind = _kind(n)
        verb = "added" if n not in a else "removed" if n not in b else "changed" if a[n] != b[n] else None
        if verb is None:
            continue
        info = (db if verb != "removed" else da).get(n)
        if info is None:
            bump(kind, verb)
        elif verb == "changed":
            old_d, new_d = da.get(n), db[n]
            changes = (diff_chart if kind == "Chart" else diff_pivot)(old_d[1], new_d[1]) if old_d else []
            label = _chart_label(new_d[1]) if kind == "Chart" else _pivot_label(new_d[1])
            lines.append(f"{label}: " + ("; ".join(changes) if changes else "formatting or other settings changed"))
        else:
            label = _chart_label(info[1]) if kind == "Chart" else _pivot_label(info[1])
            lines.append(f"{label}: {verb}")
    for (kind, verb), count in generic.items():
        lines.append(f"Macros: {verb}" if kind == "Macros" else f"{count} {kind.lower()}{'s' if count > 1 else ''}: {verb}")
    return lines
