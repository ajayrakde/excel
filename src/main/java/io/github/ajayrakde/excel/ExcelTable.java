package io.github.ajayrakde.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.LinkedHashMap;
import java.util.Map;

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
        checkCanAdd();
        for (Object value : values) ExcelSheet.validateValue(value);
        rows.add(new ArrayList<>(Arrays.asList(values)));
        return this;
    }

    /** Adds an empty row that can be populated by configured header name. */
    public NamedRow addRow() {
        book.ensureOpen();
        if (config.headers().isEmpty()) {
            throw new IllegalStateException(name + ": YAML headers are required for named values");
        }
        checkCanAdd();
        List<Object> values = new ArrayList<>(java.util.Collections.nCopies(config.address().columns(), null));
        rows.add(values);
        Map<String, Integer> indexes = new LinkedHashMap<>();
        for (int i = 0; i < config.headers().size(); i++) indexes.put(config.headers().get(i), i);
        return new NamedRow(values, Map.copyOf(indexes));
    }

    private void checkCanAdd() {
        if (config.limit() != null && rows.size() >= config.limit()) {
            throw new IllegalArgumentException(name + ": data row limit is " + config.limit());
        }
        if ((long) firstDataRow() + rows.size() >= 1_048_576) {
            throw new IllegalArgumentException(name + ": exceeds XLSX row limit");
        }
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

    public final class NamedRow {
        private final List<Object> values;
        private final Map<String, Integer> indexes;

        private NamedRow(List<Object> values, Map<String, Integer> indexes) {
            this.values = values;
            this.indexes = indexes;
        }

        public NamedRow set(String header, Object value) {
            book.ensureOpen();
            String localName = header != null && header.startsWith(name + ".")
                    ? header.substring(name.length() + 1) : header;
            Integer index = indexes.get(localName);
            if (index == null) throw new IllegalArgumentException("Unknown table header: " + header);
            ExcelSheet.validateValue(value);
            values.set(index, value);
            return this;
        }
    }
}
