package io.github.ajayrakde.excel;

import java.io.IOException;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/** YAML-configured report. Writes are buffered until save; close does not save. Not thread-safe. */
public final class ExcelReport implements AutoCloseable {
    private final ExcelBook book;
    private final Map<String, ReportSheet> sheets = new LinkedHashMap<>();

    private ExcelReport(ExcelBook book, Map<String, Map<String, ReportLayout.Field>> layout) {
        this.book = book;
        layout.forEach((name, fields) -> sheets.put(name,
                new ReportSheet(book, book.existingSheet(name), fields)));
    }

    public static ExcelReport open(String template, String layout) throws IOException {
        return open(Path.of(template), Path.of(layout));
    }

    public static ExcelReport open(Path template, Path layout) throws IOException {
        var config = ReportLayout.load(layout);
        ExcelBook book = ExcelBook.open(template);
        try {
            return new ExcelReport(book, config);
        } catch (RuntimeException error) {
            try { book.close(); } catch (IOException closeError) { error.addSuppressed(closeError); }
            throw error;
        }
    }

    public ReportSheet sheet(String name) {
        book.ensureOpen();
        ReportSheet sheet = sheets.get(name);
        if (sheet == null) throw new IllegalArgumentException("Unknown configured sheet: " + name);
        return sheet;
    }

    public void save(String output) throws IOException { save(Path.of(output)); }

    /** Validates all pending destinations before modifying the workbook or opening the output file. */
    public void save(Path output) throws IOException {
        book.ensureOpen();
        sheets.values().forEach(ReportSheet::validate);
        sheets.values().forEach(ReportSheet::apply);
        book.save(output);
    }

    @Override public void close() throws IOException { book.close(); }
}
