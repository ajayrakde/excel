package io.github.ajayrakde.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Entry point for filling an XLSX template from a YAML layout. Not thread-safe. */
public final class ExcelBook implements AutoCloseable {
    private final XSSFWorkbook workbook;
    private final Map<String, ExcelSheet> sheets = new LinkedHashMap<>();
    private boolean closed;

    private ExcelBook(XSSFWorkbook workbook, Map<String, Map<String, Layout.Field>> layout) {
        this.workbook = workbook;
        layout.forEach((name, fields) -> {
            Sheet sheet = workbook.getSheet(name);
            if (sheet == null) throw new IllegalArgumentException("Template has no sheet: " + name);
            sheets.put(name, new ExcelSheet(this, sheet, fields));
        });
    }

    public static ExcelBook open(String template, String layout) throws IOException {
        return open(Path.of(template), Path.of(layout));
    }

    public static ExcelBook open(Path template, Path layout) throws IOException {
        var config = Layout.load(layout);
        XSSFWorkbook workbook;
        try (InputStream input = Files.newInputStream(template)) {
            workbook = new XSSFWorkbook(input);
        }
        try {
            return new ExcelBook(workbook, config);
        } catch (RuntimeException error) {
            try { workbook.close(); } catch (IOException closeError) { error.addSuppressed(closeError); }
            throw error;
        }
    }

    public ExcelSheet sheet(String name) {
        ensureOpen();
        ExcelSheet sheet = sheets.get(name);
        if (sheet == null) throw new IllegalArgumentException("Unknown configured sheet: " + name);
        return sheet;
    }

    public void save(String output) throws IOException { save(Path.of(output)); }

    /** Validates every pending write before changing workbook cells or opening the output file. */
    public void save(Path output) throws IOException {
        ensureOpen();
        sheets.values().forEach(ExcelSheet::validate);
        sheets.values().forEach(ExcelSheet::apply);
        try (OutputStream stream = Files.newOutputStream(output)) {
            workbook.write(stream);
        }
    }

    void ensureOpen() {
        if (closed) throw new IllegalStateException("Book is closed");
    }

    @Override public void close() throws IOException {
        if (!closed) {
            closed = true;
            workbook.close();
        }
    }
}
