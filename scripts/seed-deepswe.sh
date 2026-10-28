#!/bin/sh
# Seed the "DeepSWE runs" workbook through the sheets CLI (doubles as a test of the agent path).
# Usage: scripts/seed-deepswe.sh [csv]   (SHEETS_URL picks the server)
set -eu
csv="${1:-$HOME/deepswe-results/clean/deepswe-runs-sheet.csv}"
sheets() { "$(dirname "$0")/../bin/sheets" "$@"; }
B="DeepSWE runs"

sheets import "$csv" --name "$B"
sheets tab "$B" rename Sheet1 Runs
sheets freeze "$B" 1 1

# Header row: bold labels reading bottom-to-top, so narrow columns can show long names.
sheets format "$B" A1:BP1 --bold
sheets format "$B" B1:BP1 --rotate
sheets height "$B" 1:1 250

# Run metadata columns.
sheets width "$B" A:A 300
sheets width "$B" B:C 84
sheets width "$B" D:D 230
sheets width "$B" E:E 96
sheets width "$B" F:F 164
sheets width "$B" G:G 112
sheets width "$B" H:Q 56
sheets width "$B" R:R 120
sheets format "$B" M2:O19 --nf 0.0%

# Pass/fail matrix: 1 pass, 0 fail, x excluded, blank not run.
sheets width "$B" S:BP 28
sheets format "$B" S2:BP19 --align center
sheets cond "$B" add S2:BP19 eq 1 --fill green
sheets cond "$B" add S2:BP19 eq 0 --fill red
sheets cond "$B" add S2:BP19 eq x --fill grey
sheets show "$B"
