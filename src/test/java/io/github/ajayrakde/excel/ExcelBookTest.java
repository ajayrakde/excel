package io.github.ajayrakde.excel;

import static org.junit.jupiter.api.Assertions.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

class ExcelBookTest {
    @TempDir Path directory;

    private Path template(boolean merge) throws Exception {
        Path template = directory.resolve("template.xlsx");
        try (var book = new XSSFWorkbook()) {
            var a = book.createSheet("A");
            a.createRow(0).createCell(0).setCellValue("Name");
            var header = a.createRow(1);
            header.createCell(0).setCellValue("#");
            header.createCell(1).setCellValue("Item name");
            header.createCell(2).setCellValue("Price");
            a.createRow(7).createCell(0).setCellValue("outside");
            book.createSheet("B");
            if (merge) a.addMergedRegion(CellRangeAddress.valueOf("A3:B3"));
            try (var output = Files.newOutputStream(template)) { book.write(output); }
        }
        return template;
    }

    private Path layout(String text) throws Exception {
        Path path = directory.resolve("layout.yml");
        Files.writeString(path, text);
        return path;
    }

    @Test void populatesMultipleSheetsPreservesHeadersAndLimitsAndSupportsRepeatedSave() throws Exception {
        Path input = template(false);
        Path config = layout("""
            sheets:
              A:
                customerName: {type: cell, range: B1}
                summary: {type: row, range: 'E1:G1'}
                prices: {type: col, range: 'H1:H10'}
                items: {type: table, range: 'A2:C2', hasHeader: true, limit: 5,
                        headers: [number, itemName, price]}
              B:
                customerName: {type: cell, range: D1}
                items: {type: table, range: 'B4:D4', hasHeader: false}
            """);
        Path output = directory.resolve("output.xlsx");
        try (var report = ExcelBook.open(input.toString(), config.toString())) {
            var a = report.sheet("A");
            a.cell("customerName").add("Ajay");
            a.row("summary").add("Total", null, 100);
            a.col("prices").add(50, 10, 5, 8, 15, 25, 30, 12, 7, 20);
            var items = a.table("items");
            assertSame(items, report.sheet("A").table("items"));
            items.addRow().set("itemName", "Notebook").set("number", 1).set("price", 50);
            items.addRow(Map.of("items.price", 10, "items.number", 2, "items.itemName", "Pen"));
            report.sheet("B").cell("customerName").add("Vijay");
            report.sheet("B").table("items").addRow(1, "Pencil", 5).addRow(2, null, 15);
            report.save(output.toString());
            items.addRow(3, "Pencil", 5).addRow(4, "Eraser", 8).addRow(5, "Ruler", 15);
            assertThrows(IllegalArgumentException.class, () -> items.addRow(6, "Overflow", 1));
            report.save(output);
        }
        try (var stream = Files.newInputStream(output); var book = new XSSFWorkbook(stream)) {
            var a = book.getSheet("A");
            assertEquals("Name", a.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Ajay", a.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Total", a.getRow(0).getCell(4).getStringCellValue());
            assertEquals(CellType.BLANK, a.getRow(0).getCell(5).getCellType());
            assertEquals(100, a.getRow(0).getCell(6).getNumericCellValue());
            assertEquals(50, a.getRow(0).getCell(7).getNumericCellValue());
            assertEquals(20, a.getRow(9).getCell(7).getNumericCellValue());
            assertEquals("Item name", a.getRow(1).getCell(1).getStringCellValue());
            assertEquals(1, a.getRow(2).getCell(0).getNumericCellValue());
            assertEquals("Notebook", a.getRow(2).getCell(1).getStringCellValue());
            assertEquals(15, a.getRow(6).getCell(2).getNumericCellValue());
            assertEquals("outside", a.getRow(7).getCell(0).getStringCellValue());
            var b = book.getSheet("B");
            assertEquals("Vijay", b.getRow(0).getCell(3).getStringCellValue());
            assertEquals(1, b.getRow(3).getCell(1).getNumericCellValue());
            assertEquals(CellType.BLANK, b.getRow(4).getCell(2).getCellType());
        }
        try (var stream = Files.newInputStream(input); var book = new XSSFWorkbook(stream)) {
            assertNull(book.getSheet("A").getRow(0).getCell(1));
        }
    }

    @Test void validatesBeforeWritingOutputAndAllowsEmptyTable() throws Exception {
        Path config = layout("""
            sheets:
              A:
                customerName: {type: cell, range: B1}
                items: {type: table, range: 'A2:C2', hasHeader: true}
            """);
        Path output = directory.resolve("output.xlsx");
        Files.writeString(output, "existing output");
        try (var report = ExcelBook.open(template(true), config)) {
            report.sheet("A").cell("customerName").add("Ajay");
            var table = report.sheet("A").table("items");
            assertThrows(IllegalArgumentException.class, () -> table.addRow(1, "too short"));
            assertThrows(IllegalArgumentException.class, () -> table.addRow(1, "invalid", new Object()));
            table.addRow(1, "overlaps merge", 5);
            assertThrows(IllegalArgumentException.class, () -> report.save(output));
            assertEquals("existing output", Files.readString(output));
        }
        try (var report = ExcelBook.open(template(false), config)) {
            report.sheet("A").table("items");
            report.save(output);
        }
        try (var stream = Files.newInputStream(output); var book = new XSSFWorkbook(stream)) {
            assertEquals("Item name", book.getSheet("A").getRow(1).getCell(1).getStringCellValue());
            assertNull(book.getSheet("A").getRow(2));
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {
        "{type: table, range: 'A2:C7', hasHeader: true}",
        "{type: table, range: 'A2:', hasHeader: true}",
        "{type: table, range: 'A2:C2', hasHeader: true, limit: 0}",
        "{type: table, range: 'A2:C2', hasHeader: true, limit: -1}",
        "{type: table, range: 'A2:C2', hasHeader: true, limit: 2.5}",
        "{type: table, range: 'A2:C2', hasHeader: true, limit: '5'}",
        "{type: table, range: 'A2:C2', hasHeader: 'true'}",
        "{type: table, range: 'A2:C2'}",
        "{type: table, range: 'A2:C2', hasHeader: false, typo: 1}",
        "{type: cell, range: 'A1:B1'}",
        "{type: cell, range: 'A1', hasHeader: true}",
        "{type: row, range: 'A1:B2'}",
        "{type: row, range: 'A1:B1', limit: 2}",
        "{type: col, range: 'A1:B2'}",
        "{type: col, range: 'A1:A2', hasHeader: false}",
        "{type: column, range: 'A1:A2'}",
        "{type: table, range: 'A2:C2', hasHeader: true, headers: [number, item]}",
        "{type: table, range: 'A2:C2', hasHeader: true, headers: [number, item, number]}",
        "{type: table, range: 'A2:C2', hasHeader: true, headers: [number, item, 3]}",
        "{type: table, range: 'A1048576:C1048576', hasHeader: true}",
        "{type: table, range: 'A1048575:C1048575', hasHeader: false, limit: 3}"
    })
    void rejectsInvalidConfiguration(String field) throws Exception {
        Path config = layout("sheets:\n  A:\n    items: " + field + "\n");
        assertThrows(IllegalArgumentException.class, () -> ExcelBook.open(directory.resolve("unused.xlsx"), config));
    }

    @Test void rejectsDuplicateYamlKeys() throws Exception {
        Path config = layout("""
            sheets:
              A:
                name: {type: cell, range: B1}
                name: {type: cell, range: C1}
            """);
        assertThrows(IllegalArgumentException.class, () -> Layout.load(config));
    }

    @Test void rejectsMissingSheetsUnknownNamesAndWrongOperationsAndClosedHandles() throws Exception {
        Path input = template(false);
        Path config = layout("""
            sheets:
              Missing:
                name: {type: cell, range: B1}
            """);
        Path missingConfig = config;
        assertThrows(IllegalArgumentException.class, () -> ExcelBook.open(input, missingConfig));
        config = layout("""
            sheets:
              A:
                name: {type: cell, range: B1}
                summary: {type: row, range: 'D1:F1'}
                prices: {type: col, range: 'G1:G3'}
                items: {type: table, range: 'A2:C2', hasHeader: true,
                        headers: [number, itemName, price]}
            """);
        var report = ExcelBook.open(input, config);
        var sheet = report.sheet("A");
        var cell = sheet.cell("name");
        var row = sheet.row("summary");
        var col = sheet.col("prices");
        var table = sheet.table("items");
        assertThrows(IllegalArgumentException.class, () -> report.sheet("B"));
        assertThrows(IllegalArgumentException.class, () -> sheet.cell("missing"));
        assertThrows(IllegalArgumentException.class, () -> sheet.cell("items"));
        assertThrows(IllegalArgumentException.class, () -> sheet.row("name"));
        assertThrows(IllegalArgumentException.class, () -> sheet.col("summary"));
        assertThrows(IllegalArgumentException.class, () -> sheet.table("name"));
        assertThrows(IllegalArgumentException.class, () -> row.add("too", "short"));
        assertThrows(IllegalArgumentException.class, () -> col.add(1, 2));
        var namedTableRow = table.addRow();
        assertThrows(IllegalArgumentException.class, () -> namedTableRow.set("unknown", 1));
        report.close();
        assertThrows(IllegalStateException.class, () -> cell.add("value"));
        assertThrows(IllegalStateException.class, () -> row.add(1, 2, 3));
        assertThrows(IllegalStateException.class, () -> col.add(1, 2, 3));
        assertThrows(IllegalStateException.class, () -> namedTableRow.set("number", 1));
        assertThrows(IllegalStateException.class, () -> table.addRow(1, "item", 5));
    }

    @Test void validatesNamedMapBeforeAddingItsRow() throws Exception {
        Path config = layout("""
            sheets:
              A:
                items: {type: table, range: 'A2:C2', hasHeader: true, limit: 1,
                        headers: [number, itemName, price]}
            """);
        Path output = directory.resolve("output.xlsx");
        try (var report = ExcelBook.open(template(false), config)) {
            var table = report.sheet("A").table("items");
            var duplicate = new LinkedHashMap<String, Object>();
            duplicate.put("number", 1);
            duplicate.put("items.number", 2);
            assertThrows(IllegalArgumentException.class, () -> table.addRow(duplicate));
            assertThrows(IllegalArgumentException.class, () -> table.addRow(Map.of("unknown", 1)));
            assertThrows(IllegalArgumentException.class, () -> table.addRow(Map.of("price", new Object())));
            table.addRow(Map.of("number", 1, "itemName", "Notebook", "price", 50));
            assertThrows(IllegalArgumentException.class, () -> table.addRow(Map.of("number", 2)));
            report.save(output);
        }
        try (var stream = Files.newInputStream(output); var book = new XSSFWorkbook(stream)) {
            assertEquals("Notebook", book.getSheet("A").getRow(2).getCell(1).getStringCellValue());
        }
    }

    @Test void acceptsLastXlsxRowAndRejectsGrowthBeyondIt() throws Exception {
        Path config = layout("""
            sheets:
              A:
                items: {type: table, range: 'A1048576:A1048576', hasHeader: false}
            """);
        try (var report = ExcelBook.open(template(false), config)) {
            var table = report.sheet("A").table("items").addRow(1);
            assertThrows(IllegalArgumentException.class, () -> table.addRow(2));
        }
    }
}
