package io.github.ajayrakde.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Rows for one configured table. */
public final class ExcelTable {
    private final ExcelBook book;
    private final String name;
    private final Layout.Field config;
    private final List<List<Object>> rows = new ArrayList<>();
    private final Map<String, Integer> headerIndexes;

    ExcelTable(ExcelBook book, String name, Layout.Field config) {
        this.book = book;
        this.name = name;
        this.config = config;
        Map<String, Integer> indexes = new LinkedHashMap<>();
        for (int i = 0; i < config.headers().size(); i++) indexes.put(config.headers().get(i), i);
        this.headerIndexes = Map.copyOf(indexes);
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

    /** Adds a row using configured header names in any order. */
    public ExcelTable addRow(Map<String, ?> namedValues) {
        book.ensureOpen();
        requireHeaders();
        if (namedValues == null) throw new IllegalArgumentException(name + ": named values are required");

        Map<Integer, Object> resolved = new LinkedHashMap<>();
        for (var entry : namedValues.entrySet()) {
            int index = resolveHeader(entry.getKey());
            if (resolved.containsKey(index)) {
                throw new IllegalArgumentException(name + ": header supplied more than once: " + entry.getKey());
            }
            ExcelSheet.validateValue(entry.getValue());
            resolved.put(index, entry.getValue());
        }

        checkCanAdd();
        List<Object> values = emptyRow();
        resolved.forEach(values::set);
        rows.add(values);
        return this;
    }

    /** Adds an empty row that can be populated by configured header name. */
    public NamedRow addRow() {
        book.ensureOpen();
        requireHeaders();
        checkCanAdd();
        List<Object> values = emptyRow();
        rows.add(values);
        return new NamedRow(values);
    }

    private void requireHeaders() {
        if (headerIndexes.isEmpty()) {
            throw new IllegalStateException(name + ": YAML headers are required for named values");
        }
    }

    private List<Object> emptyRow() {
        return new ArrayList<>(java.util.Collections.nCopies(config.address().columns(), null));
    }

    private int resolveHeader(String header) {
        String localName = header != null && header.startsWith(name + ".")
                ? header.substring(name.length() + 1) : header;
        Integer index = headerIndexes.get(localName);
        if (index == null) throw new IllegalArgumentException("Unknown table header: " + header);
        return index;
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

        private NamedRow(List<Object> values) {
            this.values = values;
        }

        public NamedRow set(String header, Object value) {
            book.ensureOpen();
            int index = resolveHeader(header);
            ExcelSheet.validateValue(value);
            values.set(index, value);
            return this;
        }
    }
}
