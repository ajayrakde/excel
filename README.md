# Excel writer

A small Java 17 library that fills named cells, rows, columns, and tables in an existing XLSX
template. Test scripts can use names from a YAML layout or direct Excel
references when a small one-off write is clearer.

## Layout

```yaml
sheets:
  A:
    customerName:
      type: cell
      range: B1
    summary:
      type: row
      range: E1:G1
    prices:
      type: col
      range: H1:H10
    items:
      type: table
      range: A2:C2
      hasHeader: true
      limit: 5
      headers: [number, itemName, price]

  B:
    customerName:
      type: cell
      range: D1
    items:
      type: table
      range: B4:D4
      hasHeader: false
      limit: 10
```

The YAML keys (`customerName`, `summary`, `prices`, and `items`) are the names
used by consumer code. A cell has one exact address. A row is a one-row range.
A column is a one-column range. A table range is its first row and fixes its
columns; it never needs an ending row.

Cells, rows, and columns also accept direct references without a YAML field:

```java
sheet.cell("B1").add("Ajay");
sheet.row("A10:E10").add("Total", 5, 100, "Paid", true);
sheet.col("G2:G6").add(50, 10, 5, 8, 15);
```

Named and direct forms can be mixed. Tables remain YAML-mapped because their
header behavior, optional limit, and logical header names are layout metadata.

- `hasHeader: true` preserves the configured row and starts data on the next row.
- `hasHeader: false` starts data on the configured row.
- `limit` is optional. It caps data rows and does not count the header.
- `headers` is an optional compact array. Its order maps names to table columns.
- Each configured worksheet must already exist in the template.

The checked-in examples cover the entire supported input surface:

- [`examples/layout.yml`](examples/layout.yml) contains multiple sheets; named
  cell, row, column, and table destinations; every cell value category; tables
  with and without physical workbook headers; tables with and without YAML
  header names; and tables with and without row limits.
- [`examples/ItemsWorkbookExample.java`](examples/ItemsWorkbookExample.java)
  fills that layout using scalar values, nulls, every supported numeric Java
  type, booleans, positional rows, locally named headers, and fully qualified
  `tableName.headerName` references.

| Possibility | Covered by |
| --- | --- |
| Named cell | `A.customerName`, `A.approved`, `A.itemCount`, `A.discount`, `A.optionalNote` |
| Named row | `A.summary` |
| Named column | `A.prices` |
| Direct cell, row, and column references | `B1`, `A10:E10`, `G2:G6` |
| Multiple worksheets | Sheets `A` and `B` |
| Physical table header | `A.items`, `B.archivedItems` |
| No physical table header | `A.rawRows`, `B.measurements` |
| YAML header names | `A.items`, `B.measurements` |
| Positional table values | `A.rawRows`, `B.archivedItems` |
| Limited table | `A.items`, `B.measurements` |
| Unlimited table | `A.rawRows`, `B.archivedItems` |

## Consumer code

```java
import io.github.ajayrakde.excel.ExcelBook;
import java.util.Map;

try (var excel = ExcelBook.open("template.xlsx", "layout.yml")) {
    var sheetA = excel.sheet("A");
    sheetA.cell("customerName").add("Ajay");
    sheetA.row("summary").add("Total", "", 100);
    sheetA.col("prices").add(50, 10, 5, 8, 15, 25, 30, 12, 7, 20);

    // Direct references are useful for writes that do not need a YAML name.
    sheetA.cell("J1").add("Direct");
    sheetA.row("J2:L2").add("Direct row", 2, true);
    sheetA.col("M1:M3").add(7, 8, 9);

    var items = sheetA.table("items");
    items.addRow()
        .set("number", 1)
        .set("itemName", "Notebook")
        .set("price", 50);
    items.addRow(Map.of(
        "items.number", 2,
        "items.itemName", "Pen",
        "items.price", 10
    ));
    items.addRow(3, "Pencil", 5);

    var sheetB = excel.sheet("B");
    sheetB.cell("customerName").add("Vijay");
    sheetB.table("items")
        .addRow(1, "Notebook", 50);

    excel.save("output.xlsx");
}
```

The public API has three concepts:

| Class | Purpose |
| --- | --- |
| `ExcelBook` | Opens the template and layout, selects sheets, saves and closes |
| `ExcelSheet` | Selects named cells, rows, columns, and tables |
| `ExcelTable` | Adds positional rows, maps, or fluent rows populated by header name |

`Layout` and `Address` are package-private implementation details.

## Behavior

- Supported cell values are strings, booleans, finite byte/short/int/long/float/
  double values, and `null`. A null clears the target cell.
- Strings beginning with `=` remain literal strings; formulas are not generated.
- Every table row must contain exactly the number of columns declared by its range.
- Every row or column input must exactly match its configured or direct range.
- Header names are unique and must match the table width. Within a selected table,
  use `itemName`; the fully qualified form `items.itemName` is also accepted.
- Numbering is explicit data supplied to `addRow`; it is not generated.
- Unknown sheets or names, incorrect types, invalid ranges, duplicate YAML keys,
  unsupported YAML properties, bad values, and row-limit overflow fail clearly.
- Pending writes are validated for merged-cell conflicts before `save` modifies
  workbook cells or opens the output file.
- Existing cell formatting is retained when a value is replaced.
- Empty tables leave the template unchanged. Cells below supplied data are not
  cleared, so start with a clean template when old data must not remain.
- Repeated calls to `table(name)` return the same table and continue appending.
  Repeated saves write buffered rows to the same locations without duplication.
- `close()` releases the workbook and does not save automatically.

## Build

Requires JDK 17 and Gradle 8.14.3:

```shell
gradle build
```

GitHub Actions runs the complete Gradle build for pushes and pull requests.
