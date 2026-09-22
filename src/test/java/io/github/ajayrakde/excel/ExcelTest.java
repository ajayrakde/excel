package io.github.ajayrakde.excel;

import static org.junit.jupiter.api.Assertions.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

class ExcelTest {
    @TempDir Path directory;

    @Test void writesAllFourShapesAcrossSheets() throws Exception {
        Path output = directory.resolve("out.xlsx");
        try (ExcelBook book = ExcelBook.create()) {
            book.sheet("January")
                .write("B2", "Ajay")
                .write("B3:D3", List.of("Item", 2, true))
                .write("F3:F5", List.of(10, 20, 30))
                .write("C5:D6", List.of(List.of("A", 100), List.of("B", 200)));
            book.sheet("February").write("C5:D6",
                List.of(List.of("C", 300), List.of("D", 400)));
            book.save(output);
        }
        try (var input = Files.newInputStream(output); var workbook = new XSSFWorkbook(input)) {
            var jan = workbook.getSheet("January");
            assertEquals("Ajay", jan.getRow(1).getCell(1).getStringCellValue());
            assertEquals("Item", jan.getRow(2).getCell(1).getStringCellValue());
            assertEquals(2, jan.getRow(2).getCell(2).getNumericCellValue());
            assertTrue(jan.getRow(2).getCell(3).getBooleanCellValue());
            assertEquals(30, jan.getRow(4).getCell(5).getNumericCellValue());
            assertEquals("A", jan.getRow(4).getCell(2).getStringCellValue());
            assertEquals(200, jan.getRow(5).getCell(3).getNumericCellValue());
            assertEquals(400, workbook.getSheet("February").getRow(5).getCell(3).getNumericCellValue());
            assertNull(jan.getRow(5).getCell(4));
        }
    }

    @Test void rejectsInvalidWritesBeforeAnyChanges() throws Exception {
        Path output = directory.resolve("out.xlsx");
        try (ExcelBook book = ExcelBook.create()) {
            var sheet = book.sheet("Data");
            sheet.write("A1:B1", List.of("original", "keep"));
            for (Object invalid : List.of(List.of(1), List.of(1, 2, 3),
                    List.of("changed", new Object()), List.of("changed", Double.NaN),
                    List.of("changed", "x".repeat(32768)))) {
                assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:B1", invalid));
            }
            assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:B2", List.of(1, 2, 3, 4)));
            assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:B2", List.of(List.of(1, 2), List.of(3))));
            assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:B1", List.of()));
            assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:B1", "scalar"));
            book.save(output);
        }
        try (var input = Files.newInputStream(output); var workbook = new XSSFWorkbook(input)) {
            assertEquals("original", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("keep", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        }
    }

    @Test void opensTemplatePreservesStyleAndRejectsMergedOverlap() throws Exception {
        Path template = directory.resolve("template.xlsx");
        Path output = directory.resolve("output.xlsx");
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");
            var row = sheet.createRow(0);
            var cell = row.createCell(0);
            cell.setCellFormula("1+1");
            var style = workbook.createCellStyle();
            style.setWrapText(true);
            cell.setCellStyle(style);
            row.createCell(1).setCellValue("clear");
            row.createCell(2).setCellValue("merged");
            sheet.addMergedRegion(CellRangeAddress.valueOf("C1:D1"));
            try (var stream = Files.newOutputStream(template)) { workbook.write(stream); }
        }
        try (var book = ExcelBook.open(template)) {
            var sheet = book.sheet("Data");
            assertThrows(IllegalArgumentException.class, () -> sheet.write("A1:D1", List.of(1, 2, 3, 4)));
            assertThrows(IllegalArgumentException.class, () -> sheet.write("D1", "hidden"));
            sheet.write("A1:B1", Arrays.asList("=literal", null));
            sheet.write("A2", 10).write("A2", null);
            book.save(output);
        }
        try (var input = Files.newInputStream(output); var workbook = new XSSFWorkbook(input)) {
            var sheet = workbook.getSheetAt(0);
            assertEquals(CellType.STRING, sheet.getRow(0).getCell(0).getCellType());
            assertEquals("=literal", sheet.getRow(0).getCell(0).getStringCellValue());
            assertTrue(sheet.getRow(0).getCell(0).getCellStyle().getWrapText());
            assertEquals(CellType.BLANK, sheet.getRow(0).getCell(1).getCellType());
            assertEquals(CellType.BLANK, sheet.getRow(1).getCell(0).getCellType());
            assertEquals(1, sheet.getNumMergedRegions());
            assertEquals("merged", sheet.getRow(0).getCell(2).getStringCellValue());
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"", "A0", "XFE1", "A1048577", "B2:A1", "A1:B2:C3", "Sheet!A1", "A:A", "1:2", "A1:"})
    void rejectsInvalidAddresses(String address) {
        assertThrows(IllegalArgumentException.class, () -> Address.parse(address));
    }

    @Test void parsesAbsoluteLowercaseAndBoundaryAddresses() {
        assertEquals(new Address(0, 0, 1, 1), Address.parse(" $a$1:$b$2 "));
        assertEquals(new Address(1048575, 16383, 1048575, 16383), Address.parse("XFD1048576"));
        assertThrows(IllegalArgumentException.class, () -> Address.parse(null));
    }

    @Test void closedWorkbookRejectsOperations() throws Exception {
        ExcelBook book = ExcelBook.create();
        ExcelSheet sheet = book.sheet("Data");
        book.close();
        book.close();
        assertThrows(IllegalStateException.class, () -> sheet.write("A1", 1));
        assertThrows(IllegalStateException.class, () -> book.sheet("Other"));
        assertThrows(IllegalStateException.class, () -> book.save(directory.resolve("closed.xlsx")));
    }
}
