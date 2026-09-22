import io.github.ajayrakde.excel.ExcelBook;
import java.nio.file.Path;

/** Complete consumer example for examples/layout.yml and an existing workbook. */
public final class ItemsWorkbookExample {
    private ItemsWorkbookExample() {}

    public static void fill(Path existingWorkbook, Path outputWorkbook) throws Exception {
        try (var book = ExcelBook.open(existingWorkbook, Path.of("examples/layout.yml"))) {
            var sheetA = book.sheet("A");

            // Named cells: text, boolean, integer, decimal, and blank.
            sheetA.set("customerName", "Ajay");
            sheetA.set("approved", true);
            sheetA.set("itemCount", 5);
            sheetA.set("discount", 7.5);
            sheetA.set("optionalNote", null);

            var items = sheetA.table("items");

            // Named values can be supplied in any order using local header names.
            items.addRow()
                .set("itemName", "Notebook")
                .set("price", 50)
                .set("number", 1);

            // Fully qualified tableName.headerName references are also accepted.
            items.addRow()
                .set("items.price", 10)
                .set("items.number", 2)
                .set("items.itemName", "Pen");

            // Positional rows remain available even when YAML headers are present.
            items.addRow(3, "Pencil", 5);

            // A table without YAML headers accepts positional rows.
            sheetA.table("rawRows")
                .addRow("text", true, null)
                .addRow((byte) 1, (short) 2, 3)
                .addRow(4L, 5.5f, 6.75d);

            var sheetB = book.sheet("B");
            sheetB.set("status", "Ready");

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
