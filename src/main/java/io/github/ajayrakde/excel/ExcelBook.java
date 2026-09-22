package io.github.ajayrakde.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns an XLSX workbook. Use in try-with-resources. Not thread-safe. */
public final class ExcelBook implements AutoCloseable {
    private final XSSFWorkbook workbook;
    private boolean closed;

    private ExcelBook(XSSFWorkbook workbook) { this.workbook = workbook; }

    public static ExcelBook create() { return new ExcelBook(new XSSFWorkbook()); }

    public static ExcelBook open(Path path) throws IOException {
        try (InputStream input = Files.newInputStream(path)) {
            return new ExcelBook(new XSSFWorkbook(input));
        }
    }

    /** Select an existing sheet, or create it if absent. */
    public ExcelSheet sheet(String name) {
        ensureOpen();
        Sheet sheet = workbook.getSheet(name);
        if (sheet == null) sheet = workbook.createSheet(name);
        return new ExcelSheet(this, sheet);
    }

    /** Saves to the supplied path, replacing an existing file. */
    public void save(Path path) throws IOException {
        ensureOpen();
        try (OutputStream output = Files.newOutputStream(path)) {
            workbook.write(output);
        }
    }

    void ensureOpen() {
        if (closed) throw new IllegalStateException("Workbook is closed");
    }

    ExcelSheet existingSheet(String name) {
        ensureOpen();
        Sheet sheet = workbook.getSheet(name);
        if (sheet == null) throw new IllegalArgumentException("Template has no sheet: " + name);
        return new ExcelSheet(this, sheet);
    }

    @Override public void close() throws IOException {
        if (!closed) {
            closed = true;
            workbook.close();
        }
    }
}
