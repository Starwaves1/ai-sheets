// RFC 4180 CSV (and TSV, for clipboard interop with Google Sheets and Excel).

export function parseDelimited(text: string, delim?: "," | "\t"): string[][] {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const d = delim ?? guessDelimiter(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let i = 0;
  let quoted = false;
  let any = false;
  while (i < text.length) {
    const ch = text[i]!;
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === "") {
      quoted = true;
      any = true;
      i++;
      continue;
    }
    if (ch === d) {
      row.push(field);
      field = "";
      any = true;
      i++;
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      any = false;
      i += ch === "\r" && text[i + 1] === "\n" ? 2 : 1;
      continue;
    }
    field += ch;
    any = true;
    i++;
  }
  if (any || field !== "" || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function guessDelimiter(text: string): "," | "\t" {
  const nl = text.indexOf("\n");
  const line = nl < 0 ? text : text.slice(0, nl);
  return (line.match(/\t/g)?.length ?? 0) > (line.match(/,/g)?.length ?? 0) ? "\t" : ",";
}

function quote(field: string, d: string): string {
  return field.includes(d) || /["\r\n]/.test(field) || (d === "," && /^\s|\s$/.test(field))
    ? `"${field.replace(/"/g, '""')}"`
    : field;
}

export function toDelimited(rows: string[][], delim: "," | "\t" = ","): string {
  return rows.map((r) => r.map((f) => quote(f, delim)).join(delim)).join("\n") + (rows.length ? "\n" : "");
}
