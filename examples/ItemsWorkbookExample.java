import io.github.ajayrakde.excel.ExcelBook;
import java.nio.file.Path;
import java.util.Map;

/** Complete consumer example for examples/layout.yml and an existing workbook. */
public final class ItemsWorkbookExample {
    private ItemsWorkbookExample() {}

    public static void fill(Path existingWorkbook, Path outputWorkbook) throws Exception {
        try (var book = ExcelBook.open(existingWorkbook, Path.of("examples/layout.yml"))) {
            var sheetA = book.sheet("A");

            // Named cells: text, boolean, integer, decimal, and blank.
            sheetA.cell("customerName").add("Ajay");
            sheetA.cell("approved").add(true);
            sheetA.cell("itemCount").add(5);
            sheetA.cell("discount").add(7.5);
            sheetA.cell("optionalNote").add(null);

            // A row and column must exactly match their configured ranges.
            sheetA.row("summary").add("Total", null, 100);
            sheetA.col("prices").add(50, 10, 5, 8, 15, 25, 30, 12, 7, 20);

            // Direct references are optional alternatives to YAML field names.
            sheetA.cell("J1").add("Direct");
            sheetA.row("J2:L2").add("Direct row", 2, true);
            sheetA.col("M1:M3").add(7, 8, 9);

            var items = sheetA.table("items");

            // Named values can be supplied in any order using local header names.
            items.addRow()
                .set("itemName", "Notebook")
                .set("price", 50)
                .set("number", 1);

            // A map also accepts fully qualified tableName.headerName references.
            items.addRow(Map.of(
                "items.price", 10,
                "items.number", 2,
                "items.itemName", "Pen"
            ));

            // Positional rows remain available even when YAML headers are present.
            items.addRow(3, "Pencil", 5);

            // A table without YAML headers accepts positional rows.
            sheetA.table("rawRows")
                .addRow("text", true, null)
                .addRow((byte) 1, (short) 2, 3)
                .addRow(4L, 5.5f, 6.75d);

            var sheetB = book.sheet("B");
            sheetB.cell("status").add("Ready");

            // Existing header, no YAML header names, and no configured row limit.
            sheetB.table("archivedItems")
                .addRow(1, "Old notebook", 25)
                .addRow(2, "Old pen", 5);

            // YAML headers work even when the workbook has no physical header row.
            sheetB.table("measurements").addRow()
                .set("label", "Length")
                .set("value", 12.75d);

            book.save(outputWorkbook);
        }
    }
}
