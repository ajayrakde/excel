# Excel writer — P0

## Simple reports with YAML

Put locations in `layout.yml`; test scripts provide only field names and data.
Each sheet has its own mappings, so the same names can be reused across sheets.

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

```java
import io.github.ajayrakde.excel.ExcelReport;

try (var excel = ExcelReport.open("template.xlsx", "layout.yml")) {
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

- A table `range` describes **only its first row** and fixes the columns. For
  `A2:C2`, every supplied row must contain exactly three values.
- `hasHeader: true` preserves that row and starts data on the next row. With
  `false`, data starts on the configured row. The flag is required for tables;
  the library does not infer or generate headings.
- Optional `limit` is a positive maximum number of **data rows**, excluding
  headers. Fewer rows are allowed. Omit it for growth up to the XLSX row limit.
- No end-only, open-column, or multi-row table range syntax is accepted here.
- Numbering is explicit in `addRow`; no automatic sequence is generated.
- Configured sheet names must already exist in the template. Unknown keys,
  duplicate YAML keys, unsupported properties, and wrong operation types fail.
- `set`/`addRow` buffer data; rows are copied and validated when added. Invalid
  additions leave the buffer unchanged. All queued writes are checked for merged
  overlaps before `save` modifies workbook cells or opens the output file.
- Repeated `table(name)` calls share the same builder. Repeated saves write the
  buffered rows at the same locations; they do not duplicate rows. Subsequent
  `addRow` calls append to the buffer. `close()` does not save automatically.
- Zero added rows leave a table unchanged. Cells beyond supplied rows are not
  cleared, so use a clean template when old data should not remain.
- Keep configured destinations distinct: overlapping mappings are not detected;
  cell writes apply before table writes. Header preservation concerns the table
  operation itself, not another mapping explicitly targeting those cells.

See `examples/layout.yml` for a complete layout. YAML is the only supported
named-layout configuration format.

A small Java 17 library for writing explicit cells and rectangular ranges in XLSX
files. Apache POI is internal. The direct address API remains available for
low-level use without configuration.

## Usage

```java
import io.github.ajayrakde.excel.ExcelBook;
import java.nio.file.Path;
import java.util.List;

try (ExcelBook book = ExcelBook.create()) {
    var sheet = book.sheet("January");
    sheet.write("B6", "Ajay");
    sheet.write("B12:D12", List.of("Product", 2, 500));
    sheet.write("G12:G14", List.of(1000, 1500, 800));
    sheet.write("A25:B27", List.of(
        List.of("Name", "Amount"),
        List.of("A", 100),
        List.of("B", 200)
    ));
    book.save(Path.of("output.xlsx"));
}
```

Use `ExcelBook.open(Path.of("template.xlsx"))` to edit a template. `sheet(name)`
selects an existing sheet or creates it. `save(path)` replaces an existing output
file. Close the workbook with try-with-resources; instances are not thread-safe.

## P0 behavior

| Input | Behavior |
| --- | --- |
| Scalar | Exactly one cell |
| Flat `List<?>` | One row left-to-right or one column top-to-bottom |
| List of row lists | Rectangular matrix, including one-row/one-column matrices |
| Size mismatch, empty list, ragged matrix | Rejected before writing |
| `null` | Clears the destination cell; use `Arrays.asList` for null list elements |
| Existing cell | Replaces value/formula, retains formatting |
| Merged overlap | Rejected before writing, including the merge's top-left cell |
| String | Literal text, including text starting with `=`; max 32767 characters |
| Boolean | Excel boolean |
| Byte, Short, Integer, Long, Float, Double | Excel number; must be finite |
| Other types | Rejected; convert dates/identifiers to text explicitly in P0 |

Ranges are inclusive and must be ordered from top-left to bottom-right. A1
addresses accept lowercase and `$` markers; sheet-qualified addresses, whole
rows/columns, and disjoint ranges are not supported. Limits are XLSX limits:
XFD1048576. Numeric values use Excel's double precision (roughly 15 significant
digits); use strings for long identifiers or values needing exact decimal text.

Validation of shape, values, and merged overlaps completes before mutation.
This is not a transactional guarantee against I/O failure, resource exhaustion,
or concurrent mutation of input lists. No broadcasting, truncation, implicit
range expansion, or clearing outside the destination occurs.

Search, relative positioning, formatting APIs, formula generation, merge/unmerge,
object-field mapping, and pattern filling are outside P0.

## Build and test

Requires JDK 17 and Gradle 8.14.3. Run `gradle build` (or `gradle test` for tests only).
GitHub Actions installs this Gradle version and runs the build on pushes
and pull requests. Tests reopen generated workbooks to verify actual saved cell
types, values, shared layouts, template preservation, and validation behavior.
