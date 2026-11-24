# Tablient Spreadsheet Diff

**See what changed in an Excel workbook, right in the pull request.**

GitHub shows a changed `.xlsx` as "Binary file not shown". This Action posts one comment on the pull request that shows,
sheet by sheet, which values, formulas and number formats changed. It updates the same comment on later pushes. An optional
AI summary lists the changes in plain English.

````diff
@@ -3,8 +3,8 @@
 #|A|B|C
 1|Line|Q1|Q2
-2|Units|1200|1260
+2|Units|1250|1337.5
 3|Price|12|12

 ```formulas
-C2 = B2*1.05
+C2 = B2*1.07
 ```
````

A formula change and a typed-over value read as what they are, with their cell addresses, instead of a wall of XML.

## Use it

Add `.github/workflows/xlsx-diff.yml` to the repository that holds your workbooks:

```yaml
name: Spreadsheet diff
on:
  pull_request:
    paths: ["**/*.xlsx", "**/*.xlsm"]
permissions:
  contents: read
  pull-requests: write
jobs:
  diff:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0          # the old side is read from the merge base, so a shallow clone is not enough
      - uses: tablient/spreadsheet-diff@v1
        # with:
        #   summarize: "true"
        #   anthropic-api-key: ${{ secrets.ANTHROPIC_API_KEY }}
```

## Inputs

| Input | Default | What it does |
|---|---|---|
| `github-token` | `${{ github.token }}` | Posts and updates the comment. Needs `pull-requests: write`. |
| `summarize` | `"false"` | `"true"` adds a short AI summary of the changes. |
| `anthropic-api-key` | none | Needed only for the summary. |
| `model` | `claude-haiku-5-5` | Model for the summary. |
| `cellprose-source` | `cellprose>=0.3.0,<0.4` | pip requirement for [Cellprose](https://github.com/cellprose/cellprose), the open format each workbook is turned into before diffing. |

## Privacy

Workbooks are read on the runner. Nothing leaves it except the optional summary, which sends the diff text (not the
workbooks) to the model you chose with your own key.

## Limits

- The diff compares the Cellprose text of each sheet, not a cell-by-cell model. Rows are matched by content, so an inserted
  row shows as one added row. References to shifted rows are compared after the shift.
- Charts and pivot tables are summarised (type, title, series ranges; pivot source, fields and aggregations), not drawn.
  Macros and images are only reported as added, removed or changed.
- One comment holds about 60,000 characters; longer diffs are cut short and say so.
- Values of formula cells come from the file as last saved. Files written by code without cached values show `?`.

## Run it locally

```
pip install -e .        # from this folder; installs cellprose from PyPI
tablient-diffs files before.xlsx after.xlsx
tablient-diffs pr --base origin/main --head HEAD --dry-run
```

## Development

`pip install -e ".[dev]" && pytest`.

## Licence

Source-available under the [PolyForm Shield License 1.0.0](LICENSE.md). You can use this Action, read and change the
code, and run it in your own repositories, including commercial ones. You cannot use it to build or offer a product that
competes with it or with Tablient. The workbook-to-text step is [Cellprose](https://github.com/cellprose/cellprose),
which is open source under Apache-2.0.

Made by [Tablient](https://tablient.com).
