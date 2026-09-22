# Excel writer

A small Java 17 library that fills named cells and tables in an existing XLSX
template. Test scripts use names from a YAML layout and never handle cell
coordinates in Java.

## Layout

```yaml
sheets:
  A:
    customerName:
      type: cell
      range: B1
    items:
      type: table
      range: A2:C2
      hasHeader: true
      limit: 5

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

For a cell, `range` is its exact address. For a table, `range` is its first row
and fixes the number of columns. A table never needs an ending row.

- `hasHeader: true` preserves the configured row and starts data on the next row.
- `hasHeader: false` starts data on the configured row.
- `limit` is optional. It caps data rows and does not count the header.
- Each configured worksheet must already exist in the template.

See [`examples/layout.yml`](examples/layout.yml) for a complete example.

## Consumer code

```java
import io.github.ajayrakde.excel.ExcelBook;

try (var excel = ExcelBook.open("template.xlsx", "layout.yml")) {
    var sheetA = excel.sheet("A");
    sheetA.set("customerName", "Ajay");
    sheetA.table("items")
        .addRow(1, "Notebook", 50)
        .addRow(2, "Pen", 10)
        .addRow(3, "Pencil", 5)
        .addRow(4, "Eraser", 8)
        .addRow(5, "Ruler", 15);

    var sheetB = excel.sheet("B");
    sheetB.set("customerName", "Vijay");
    sheetB.table("items")
        .addRow(1, "Notebook", 50);

    excel.save("output.xlsx");
}
```

The public API has three concepts:

| Class | Purpose |
| --- | --- |
| `ExcelBook` | Opens the template and layout, selects sheets, saves and closes |
| `ExcelSheet` | Sets a named cell and selects a named table |
| `ExcelTable` | Adds rows to the selected table |

`Layout` and `Address` are package-private implementation details.

## Behavior

- Supported cell values are strings, booleans, finite byte/short/int/long/float/
  double values, and `null`. A null clears the target cell.
- Strings beginning with `=` remain literal strings; formulas are not generated.
- Every table row must contain exactly the number of columns declared by its range.
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
