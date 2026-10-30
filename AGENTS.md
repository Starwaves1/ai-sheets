# Sheets: notes for agents

A self-hosted spreadsheet at http://192.168.1.5:8790/. Garrett keeps workbooks open in a browser; anything you write shows up in his tab within a second, highlighted, with your name in the activity line. Use the `sheets` CLI (on PATH, source in `src/cli/main.ts`). Run `sheets help` for every flag.

## Flags: what Garrett wants you to look at

Garrett marks cells in the browser with the orange tag ("flag for Claude") and a note, then tells you "check my flags". It is an async note on a cell, not a chat. Work through them like this:

```sh
sheets flags                               # every open flag, all workbooks: book, cell, value, note, who, when
sheets get "DeepSWE runs" 'Runs!M2:O19'    # look at the cells and their neighbours
sheets resolve "DeepSWE runs" 'Runs!M5' "Rerun on the fixed kernel: 0.44, updated M5."
sheets flags --all                         # resolved ones too, with replies
```

`resolve` marks the flag done and attaches your reply. Both stay on the cell (green marker) until Garrett dismisses them, so say what you found or changed in one or two sentences. You can leave a flag for him too: `sheets flag <book> B4 "note"` (`--remove` takes one off). Flags live on the cells: they move with inserted or deleted rows, copy and paste, and fill.


```sh
sheets ls                                  # id, name, tabs with used range, version, last edit
sheets show "DeepSWE runs"                 # tabs, sizes, frozen panes, browser URL
sheets get "DeepSWE runs" A1:R5            # displayed values, as below
sheets get "DeepSWE runs" 'Runs!S2:BP19' --csv
sheets get budget B2:B9 --formulas         # raw inputs (formulas) instead of values
sheets get budget A1:C3 --json             # {"values": [[...]]}, numbers as numbers
```

`<book>` is an id, the exact name, or a unique name prefix (case-insensitive). With no range, `get` returns the used range of the first tab. Default output is tab-separated: a header line, column letters, then one line per row starting with the row number:

```
# Budget / Sheet1!A1:B3 (v12)
	A	B
1	item	cost
2	tea	3
```

## Write

```sh
sheets set budget B4 '=SUM(B2:B3)'         # one cell; formulas start with =
sheets set budget C2:C9 '=B2*1.2'          # one value fills a range; relative refs shift per row
sheets set budget A10 --csv rows.csv       # a block, anchored at A10 (--tsv and - for stdin work too)
sheets set budget A1 --json '[["a",1],["b",2]]'
sheets clear budget D2:D9                  # values only; add --formats to drop formatting too
sheets fill runs A2:A4 A5:A40              # continue a pattern, like dragging the fill handle
```

`fill` continues what the source shows: 1,2,3 → 4,5,6; 10,7 → 4,1; `Item 1` → `Item 2`; `run-01` → `run-02`; Mon → Tue; Jan → Feb; anything else repeats. Formulas shift their relative references, formats come along, and conditional rules covering the source grow to cover the target. The target can include the source (A2:A40) or sit beside it in any direction; filling up or left runs the series backwards. `--copy` repeats instead of counting (and counts up from a single number, which otherwise copies).

- Write only the cells you mean to change. Edits are per cell, so you and Garrett can work on the same sheet at once; rewriting a whole range you only read can overwrite his newer edits.
- Text that starts with `=` but is not a formula: prefix an apostrophe (`'=not a formula`).
- Set `SHEETS_ACTOR=<your name>` so Garrett can tell who changed what (default `agent`).
- `sheets watch <book>` prints every change as it happens, including Garrett's.

## Format and structure

```sh
sheets format runs A1:R1 --bold --fill grey        # --italic --strike --color C --align left|center|right
sheets format runs M2:O19 --nf 0.0%                # number formats: 0  0.00  #,##0  0%  $#,##0.00  0.00E+00
sheets cond runs add S2:BP19 eq 1 --fill green     # conditional fill; ops: eq ne gt gte lt lte between contains empty notEmpty
sheets cond runs                                   # list rules; `cond runs rm 2`, `cond runs clear`
sheets freeze runs 1 1                             # rows, columns
sheets width runs S:BP 28                          # px; `height runs 1:1 200`
sheets insert runs 5:6                             # two rows before row 5; columns: `insert runs C:C`
sheets delete runs C:D                             # formulas elsewhere are rewritten, like any spreadsheet
sheets move runs 5:6 --to 2                        # rows 5-6 become rows 2-3; `move runs C:D --to H` for columns
sheets hide runs R:R                               # unhide the same way
sheets sort runs M --desc                          # shared view sort; `sort runs --off`
sheets filter runs B --hide invalid,smoke          # or: `filter runs K gt 10`; `filter runs --clear`
sheets tab runs add Notes                          # also: rename <old> <new>, rm <name> --yes, move <name> <index>
sheets import results.csv --name "Results"         # new workbook; `--book runs --sheet Extra` adds a tab instead
sheets export runs Runs > runs.csv
```

Colors are theme-aware names: `green red amber blue violet grey` (or `#rrggbb`).

Sort and filter are a shared view: they change what everyone sees, never where data is stored. A1 addresses always mean stored rows, so `B5` is the same cell whatever the sort.

## HTTP API

Everything the CLI does goes through JSON over HTTP; send `X-Actor: <name>`.

| Method and path | Does |
|---|---|
| `GET /api/books` | list workbooks |
| `POST /api/books` | `{"name"}` creates; a CSV/TSV body with `?name=` imports |
| `GET /api/books/<book>` | summary: sheets, used ranges, version |
| `GET /api/books/<book>/range?ref=Runs!A1:C9&as=json\|csv\|tsv&show=values\|inputs\|display\|both` | read |
| `PUT /api/books/<book>/range?ref=B2` | write `{"values": [[...]]}` or a CSV/TSV body |
| `DELETE /api/books/<book>/range?ref=B2:C9[&formats=1]` | clear |
| `POST /api/books/<book>/format?ref=A1:C1` | `{"b": true, "fill": "green", "nf": "0.0%"}`; `null` removes a key |
| `POST /api/books/<book>/ops` | `{"ops": [...], "base": <version>}`: raw ops, see `src/core/ops.ts` |
| `GET /api/books/<book>/events` | server-sent events, one `change` per edit |
| `GET /api/books/<book>/csv?sheet=Runs` | export |

Moves keep everything attached to the rows or columns (values, formats, flags, sizes, hidden state), and references follow them: `=A5` becomes `=A2` when row 5 moves to row 2; a range keeps covering the same rows. Raw ops use 0-based row/col indices and sheet ids from the summary, for example `{"t":"insert","s":"<sheet id>","axis":"row","at":4,"n":2}`. The `Op` type in `src/core/ops.ts` lists them all.

## Working on the code

- `npm run build` bundles server, CLI and browser app into `dist/`; `npm test` runs core and API tests; `SHEETS_URL=http://192.168.1.5:8790 npm run e2e` checks the live path and idle cost in headless Chromium; `npm run mobile` checks phones and an iPad (touch emulation).
- The service is the systemd user unit `sheets.service` (`systemctl --user restart sheets` after a build). Data lives in `data/sheets.db` (git-ignored).
- `src/core` is pure and shared by server and browser: A1 utilities, the document model, ops (apply, inverse, transform), the formula engine, display formats. Add a formula function in `src/core/formula/functions.ts` with a test in `test/core.test.ts`.
- `src/server` is the HTTP API, SSE fan-out and SQLite op log; `src/client` is the browser app (no framework; `grid.ts` is the virtualized grid).
- Keep the browser idle when nothing changes: no timers, polling, or looping animations. `npm run e2e` fails if the page does work while idle.
