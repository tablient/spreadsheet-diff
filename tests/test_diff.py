import openpyxl

from tablient_diffs.diff import diff_workbooks
from tablient_diffs.render import MARKER, render


def book(tmp_path, name, rate, extra_sheet=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"
    ws.append(["Item", "Units", "Price"])
    ws.append(["Widget", 10, rate])
    ws["D1"], ws["D2"] = "Total", "=B2*C2"
    if extra_sheet:
        wb.create_sheet("Notes")["A1"] = "hello"
    p = tmp_path / name
    wb.save(p)
    return p.read_bytes()


def test_changed_value_shows_in_diff(tmp_path):
    fd = diff_workbooks("m.xlsx", book(tmp_path, "a.xlsx", 5), book(tmp_path, "b.xlsx", 7))
    assert fd.status == "modified" and [s.name for s in fd.sheets] == ["Model"]
    assert "-" in fd.sheets[0].diff and "7" in fd.sheets[0].diff


def test_identical_workbooks_report_nothing(tmp_path):
    a = book(tmp_path, "a.xlsx", 5)
    fd = diff_workbooks("m.xlsx", a, a)
    assert not fd.changed
    assert "No cell, formula or format changes" in render([fd])


def test_added_sheet_and_new_file(tmp_path):
    fd = diff_workbooks("m.xlsx", book(tmp_path, "a.xlsx", 5), book(tmp_path, "b.xlsx", 5, extra_sheet=True))
    assert any(s.name == "Notes" and s.status == "added" for s in fd.sheets)
    new = diff_workbooks("n.xlsx", None, book(tmp_path, "c.xlsx", 5))
    assert new.status == "added" and new.sheets[0].status == "added"


def test_corrupt_file_is_reported_not_raised():
    fd = diff_workbooks("bad.xlsx", None, b"not a zip")
    assert fd.status == "error" and "Could not compare" in render([fd])


def test_render_has_marker_and_summary(tmp_path):
    fd = diff_workbooks("m.xlsx", book(tmp_path, "a.xlsx", 5), book(tmp_path, "b.xlsx", 7))
    out = render([fd], summary="- price changed")
    assert out.startswith(MARKER) and "price changed" in out and "<details>" in out


def test_inserted_row_is_one_change(tmp_path):
    def make(name, rows):
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in rows:
            ws.append(r)
        p = tmp_path / name
        wb.save(p)
        return p.read_bytes()

    base = [["a", 1], ["b", 2], ["c", 3], ["d", 4]]
    new = base[:1] + [["new", 9]] + base[1:]
    fd = diff_workbooks("m.xlsx", make("a.xlsx", base), make("b.xlsx", new))
    s = fd.sheets[0]
    rows = [l for l in s.diff.splitlines() if l[:1] in "+-" and l[1:2].isdigit()]
    assert rows == ["+2|new|9"], s.diff


def test_inserted_row_does_not_mark_shifted_formulas(tmp_path):
    def make(name, n_rows, insert_at=None):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Item", "Units", "Price", "Total"])
        rows = list(range(2, 2 + n_rows))
        r = 2
        for i in range(n_rows):
            if insert_at == i:
                ws.append(["Gadget", 3, 4.5, f"=B{r}*C{r}"])
                r += 1
            ws.append([f"Item {i}", i + 1, 2.5, f"=B{r}*C{r}"])
            r += 1
        ws.append(["Sum", None, None, f"=SUM(D2:D{r - 1})"])
        p = tmp_path / name
        wb.save(p)
        return p.read_bytes()

    fd = diff_workbooks("m.xlsx", make("a.xlsx", 8), make("b.xlsx", 8, insert_at=3))
    changed = [l for l in fd.sheets[0].diff.splitlines() if l[:1] in "+-" and not l.startswith(("+++", "---"))]
    # the new row, its formula, and the grown region heading / SUM range: a handful, not every formula below
    assert len(changed) <= 8, "\n".join(changed)
    assert any("Gadget" in l for l in changed)


def test_identical_bytes_short_circuit():
    fd = diff_workbooks("m.xlsx", b"same bytes, not even a workbook", b"same bytes, not even a workbook")
    assert not fd.changed and fd.status == "modified"


def test_real_formula_edit_is_still_shown_alongside_an_inserted_row(tmp_path):
    def make(name, rows, total):
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in rows:
            ws.append(r)
        ws.append(["Sum", total])
        p = tmp_path / name
        wb.save(p)
        return p.read_bytes()

    base = [["a", 1], ["b", 2], ["c", 3]]
    old = make("a.xlsx", base, "=SUM(B1:B3)")
    new = make("b.xlsx", base[:1] + [["x", 9]] + base[1:], "=AVERAGE(B1:B4)")   # row inserted AND formula changed
    diff = diff_workbooks("m.xlsx", old, new).sheets[0].diff
    assert "AVERAGE" in diff and "SUM" in diff


def test_pr_reads_old_side_from_merge_base(tmp_path, monkeypatch, capsys):
    import subprocess

    from tablient_diffs.cli import main

    repo = tmp_path / "repo"
    repo.mkdir()
    monkeypatch.chdir(repo)
    monkeypatch.delenv("GITHUB_EVENT_PATH", raising=False)

    def g(*args):
        subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@t", *args], check=True, capture_output=True)

    g("init", "-q", "-b", "main")
    (repo / "m.xlsx").write_bytes(book(tmp_path, "a.xlsx", 5))
    g("add", "m.xlsx")
    g("commit", "-qm", "base")
    g("checkout", "-qb", "pr")
    (repo / "m.xlsx").write_bytes(book(tmp_path, "b.xlsx", 7))
    g("commit", "-qam", "pr: price 7")
    g("checkout", "-q", "main")
    (repo / "m.xlsx").write_bytes(book(tmp_path, "c.xlsx", 5, extra_sheet=True))
    g("commit", "-qam", "main moved on: add Notes")

    assert main(["pr", "--base", "main", "--head", "pr", "--dry-run"]) == 0
    out = capsys.readouterr().out
    assert "Model" in out and "Notes" not in out


def test_metadata_only_change_reports_nothing(tmp_path):
    a = book(tmp_path, "a.xlsx", 5)
    wb = openpyxl.load_workbook(tmp_path / "a.xlsx")
    wb.properties.creator = "someone else"
    wb.save(tmp_path / "b.xlsx")
    b = (tmp_path / "b.xlsx").read_bytes()
    assert a != b
    fd = diff_workbooks("m.xlsx", a, b)
    assert not fd.changed and fd.sheets == []
    assert "No cell, formula or format changes" in render([fd])


def test_render_respects_the_comment_size_limit():
    from tablient_diffs.diff import FileDiff, SheetChange
    from tablient_diffs.render import MAX_COMMENT

    big = "\n".join(f"+|{i}|" + "x" * 150 for i in range(150))
    files = [FileDiff(f"f{n}.xlsx", "modified", [SheetChange(f"S{k}", "changed", 150, 0, big) for k in range(3)])
             for n in range(20)]
    out = render(files, summary="s" * 20_000)
    assert len(out) <= MAX_COMMENT
    shown = out.count("#### ")
    assert 0 < shown < 20
    assert out.count("more file(s) omitted") == 1
    assert f"_{20 - shown} more file(s) omitted" in out
    assert out.rstrip().endswith("</sub>")


def test_sheet_named_like_an_internal_key_does_not_collide():
    from tablient_diffs.diff import split_sheets

    text = "---\nsheetmark: 0.2\n---\n# Sheet: (front matter)\n|#|A|\n|1|x|\n# Sheet: (defined names)\n|#|A|\n|1|y|\n# Names\nN = A1\n"
    parts = split_sheets(text)
    assert parts["(front matter)"][1:] == ["|#|A|", "|1|x|"]
    assert parts["(defined names)"][1:] == ["|#|A|", "|1|y|"]
    assert len(parts) == 4


def test_chart_macro_and_pivot_parts_are_reported(tmp_path):
    import io
    import zipfile

    def with_parts(data, parts):
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as z:
            for n in zin.namelist():
                z.writestr(n, zin.read(n))
            for n, c in parts.items():
                z.writestr(n, c)
        return out.getvalue()

    base = book(tmp_path, "a.xlsx", 5)
    old = with_parts(base, {"xl/charts/chart1.xml": "<c/>", "xl/vbaProject.bin": b"v1"})
    new = with_parts(base, {"xl/charts/chart1.xml": "<c x='1'/>", "xl/vbaProject.bin": b"v1",
                            "xl/pivotTables/pivotTable1.xml": "<p/>"})
    fd = diff_workbooks("m.xlsx", old, new)
    assert not fd.sheets and fd.changed
    assert "1 chart: changed" in fd.parts and "1 pivot table: added" in fd.parts
    assert not any("Macros" in p for p in fd.parts)
    assert "Charts, pivot tables and other parts:" in render([fd])


def _sheet_book(rows, other=None, insert_at=None, blank_at=None):
    """Workbook with a Data sheet (a, b, =b*2 per row) and optionally a Summary sheet pointing at Data."""
    import io
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    r = 1
    for i in range(1, rows + 1):
        if blank_at == i:
            r += 1
        ws.cell(r, 1, f"item{i}")
        ws.cell(r, 2, i * 10)
        ws.cell(r, 3, f"=B{r}*2")
        r += 1
    if other:
        wb.create_sheet("Summary")["A1"] = f"=Data!B{other}"
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_cross_sheet_reference_follows_inserted_rows():
    # Data row 3 becomes row 4 after a blank row is inserted above it; Summary's formula points at it.
    old = _sheet_book(6, other=3)
    new = _sheet_book(6, other=4, blank_at=3)
    fd = diff_workbooks("m.xlsx", old, new)
    assert "Summary" not in [s.name for s in fd.sheets], [s.diff for s in fd.sheets]


def test_empty_row_does_not_split_the_diff():
    fd = diff_workbooks("m.xlsx", _sheet_book(10), _sheet_book(10, blank_at=6))
    rows = [l for s in fd.sheets for l in s.diff.splitlines() if l[:1] in "+-" and l[1:2].isdigit()]
    assert rows == [], fd.sheets[0].diff
    assert sum(s.added + s.removed for s in fd.sheets) <= 4, fd.sheets[0].diff


def test_unreadable_file_message_names_the_side(tmp_path):
    good = book(tmp_path, "a.xlsx", 5)
    fd = diff_workbooks("m.xlsx", good, b"not a zip")
    assert fd.status == "error" and "after version could not be read" in fd.error
    assert "damaged" in fd.error
    same = diff_workbooks("m.xlsx", b"not a zip", b"not a zip")
    assert same.status == "modified" and not same.changed


CHART = """<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:plotArea><c:{kind}Chart><c:barDir val="col"/><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
<c:cat><c:strRef><c:f>Data!$A$2:$A$10</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:{end}</c:f></c:numRef></c:val></c:ser></c:{kind}Chart>
</c:plotArea></c:chart></c:chartSpace>"""
PIVOT = """<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Sales">
<location ref="A3:C9"/><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField fld="{fld}" subtotal="{agg}"/></dataFields></pivotTableDefinition>"""
PIVOT_RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="x/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>"""
CACHE = """<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cacheSource type="worksheet"><worksheetSource ref="A1:C50" sheet="Data"/></cacheSource>
<cacheFields count="3"><cacheField name="Region"/><cacheField name="Units"/><cacheField name="Revenue"/></cacheFields></pivotCacheDefinition>"""


def _with_parts(data, parts):
    import io
    import zipfile
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as z:
        for n in zin.namelist():
            z.writestr(n, zin.read(n))
        for n, c in parts.items():
            z.writestr(n, c)
    return out.getvalue()


def test_chart_changes_are_described(tmp_path):
    base = book(tmp_path, "a.xlsx", 5)
    old = _with_parts(base, {"xl/charts/chart1.xml": CHART.format(title="Sales", kind="bar", end="$B$10")})
    new = _with_parts(base, {"xl/charts/chart1.xml": CHART.format(title="Revenue", kind="bar", end="$B$20")})
    fd = diff_workbooks("m.xlsx", old, new)
    (line,) = fd.parts
    assert "column chart 'Revenue' with 1 series" in line
    assert "title 'Sales' -> 'Revenue'" in line
    assert "Data!$B$2:$B$10 -> Data!$B$2:$B$20" in line
    added = diff_workbooks("m.xlsx", base, old)
    assert added.parts == ["column chart 'Sales' with 1 series: added"]


def test_pivot_changes_are_described(tmp_path):
    base = book(tmp_path, "a.xlsx", 5)
    shared = {"xl/pivotTables/_rels/pivotTable1.xml.rels": PIVOT_RELS, "xl/pivotCache/pivotCacheDefinition1.xml": CACHE}
    old = _with_parts(base, {**shared, "xl/pivotTables/pivotTable1.xml": PIVOT.format(fld=1, agg="sum")})
    new = _with_parts(base, {**shared, "xl/pivotTables/pivotTable1.xml": PIVOT.format(fld=2, agg="average")})
    fd = diff_workbooks("m.xlsx", old, new)
    (line,) = fd.parts
    assert "pivot table 'Sales' from Data!A1:C50" in line
    assert "values sum of Units -> average of Revenue" in line
