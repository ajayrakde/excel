package io.github.ajayrakde.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/** Rows for one configured table. */
public final class ExcelTable {
    private final ExcelBook book;
    private final String name;
    private final Layout.Field config;
    private final List<List<Object>> rows = new ArrayList<>();

    ExcelTable(ExcelBook book, String name, Layout.Field config) {
        this.book = book;
        this.name = name;
        this.config = config;
    }

    public ExcelTable addRow(Object... values) {
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

    private Address destination() {
        return new Address(firstDataRow(), config.address().firstColumn(),
                firstDataRow() + rows.size() - 1, config.address().lastColumn());
    }

    void validate(ExcelSheet sheet) {
        if (!rows.isEmpty()) sheet.validateRange(destination(), rows);
    }

    void apply(ExcelSheet sheet) {
        if (!rows.isEmpty()) sheet.write(destination(), rows);
    }
}
