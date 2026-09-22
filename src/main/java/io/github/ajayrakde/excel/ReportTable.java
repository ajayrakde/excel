package io.github.ajayrakde.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;

/** Collects data rows with fixed width and an optional maximum row count. */
public final class ReportTable {
    private final ExcelBook book;
    private final String name;
    private final ReportLayout.Field config;
    private final List<List<Object>> rows = new ArrayList<>();

    ReportTable(ExcelBook book, String name, ReportLayout.Field config) {
        this.book = book;
        this.name = name;
        this.config = config;
    }

    public ReportTable addRow(Object... values) {
        book.ensureOpen();
        if (values == null || values.length != config.address().columns()) {
            throw new IllegalArgumentException(name + ": each row requires " + config.address().columns() + " values");
        }
        if (config.limit() != null && rows.size() >= config.limit()) {
            throw new IllegalArgumentException(name + ": data row limit is " + config.limit());
        }
        if ((long) firstDataRow() + rows.size() >= 1_048_576) {
            throw new IllegalArgumentException(name + ": exceeds XLSX row limit");
        }
        for (Object value : values) ExcelSheet.validateValue(value);
        rows.add(new ArrayList<>(Arrays.asList(values)));
        return this;
    }

    private int firstDataRow() {
        return config.address().firstRow() + (config.hasHeader() ? 1 : 0);
    }

    private String destination() {
        return new CellRangeAddress(firstDataRow(), firstDataRow() + rows.size() - 1,
                config.address().firstColumn(), config.address().lastColumn()).formatAsString();
    }

    void validate(ExcelSheet sheet) {
        if (!rows.isEmpty()) sheet.validate(destination(), rows);
    }

    void apply(ExcelSheet sheet) {
        if (!rows.isEmpty()) sheet.write(destination(), rows);
    }
}
