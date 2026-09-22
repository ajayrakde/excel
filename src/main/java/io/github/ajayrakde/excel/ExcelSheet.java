package io.github.ajayrakde.excel;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/** A configured worksheet containing named cells, rows, columns, and tables. */
public final class ExcelSheet {
    private final ExcelBook book;
    private final Sheet sheet;
    private final Map<String, Layout.Field> fields;
    private final Map<String, Object> values = new LinkedHashMap<>();
    private final Map<String, ExcelTable> tables = new LinkedHashMap<>();

    ExcelSheet(ExcelBook book, Sheet sheet, Map<String, Layout.Field> fields) {
        this.book = book;
        this.sheet = sheet;
        this.fields = fields;
    }

    /** Selects a named cell. */
    public CellInput cell(String name) {
        return new CellInput(name, field(name, "cell"));
    }

    /** Selects a named horizontal range. */
    public RowInput row(String name) {
        return new RowInput(name, field(name, "row"));
    }

    /** Selects a named vertical range. */
    public ColInput col(String name) {
        return new ColInput(name, field(name, "col"));
    }

    /** Returns the same table for repeated calls, so rows continue appending. */
    public ExcelTable table(String name) {
        Layout.Field config = field(name, "table");
        return tables.computeIfAbsent(name, ignored -> new ExcelTable(book, name, config));
    }

    private Layout.Field field(String name, String expected) {
        book.ensureOpen();
        Layout.Field field = name == null ? null : fields.get(name);
        if (field == null) throw new IllegalArgumentException("Unknown field: " + name);
        if (!field.type().equals(expected)) {
            throw new IllegalArgumentException(name + " is a " + field.type() + ", not a " + expected);
        }
        return field;
    }

    void validate() {
        values.forEach((name, value) -> validateRange(fields.get(name).address(), value));
        tables.values().forEach(table -> table.validate(this));
    }

    void apply() {
        values.forEach((name, value) -> write(fields.get(name).address(), value));
        tables.values().forEach(table -> table.apply(this));
    }

    void validateRange(Address address, Object data) {
        for (CellRangeAddress merged : sheet.getMergedRegions()) {
            if (address.firstRow() <= merged.getLastRow() && address.lastRow() >= merged.getFirstRow()
                    && address.firstColumn() <= merged.getLastColumn()
                    && address.lastColumn() >= merged.getFirstColumn()) {
                throw new IllegalArgumentException("Destination overlaps merged cells " + merged.formatAsString());
            }
        }
        if (data instanceof List<?> rows) {
            if (rows.size() != address.rows()) throw shape(address);
            for (Object item : rows) {
                if (!(item instanceof List<?> row) || row.size() != address.columns()) throw shape(address);
                row.forEach(ExcelSheet::validateValue);
            }
        } else {
            if (address.rows() != 1 || address.columns() != 1) throw shape(address);
            validateValue(data);
        }
    }

    void write(Address address, Object data) {
        for (int r = 0; r < address.rows(); r++) {
            org.apache.poi.ss.usermodel.Row row = sheet.getRow(address.firstRow() + r);
            if (row == null) row = sheet.createRow(address.firstRow() + r);
            for (int c = 0; c < address.columns(); c++) {
                org.apache.poi.ss.usermodel.Cell cell = row.getCell(address.firstColumn() + c);
                if (cell == null) cell = row.createCell(address.firstColumn() + c);
                Object value = data instanceof List<?> rows ? ((List<?>) rows.get(r)).get(c) : data;
                cell.setBlank();
                if (value instanceof String text) cell.setCellValue(text);
                else if (value instanceof Boolean bool) cell.setCellValue(bool);
                else if (value instanceof Number number) cell.setCellValue(number.doubleValue());
            }
        }
    }

    static void validateValue(Object value) {
        if (value == null || value instanceof Boolean) return;
        if (value instanceof String text) {
            if (text.length() > 32_767) throw new IllegalArgumentException("Cell text exceeds 32767 characters");
            return;
        }
        if (value instanceof Byte || value instanceof Short || value instanceof Integer
                || value instanceof Long || value instanceof Float || value instanceof Double) {
            if (!Double.isFinite(((Number) value).doubleValue())) {
                throw new IllegalArgumentException("Numeric values must be finite");
            }
            return;
        }
        throw new IllegalArgumentException("Unsupported cell value type: " + value.getClass().getName());
    }

    private static IllegalArgumentException shape(Address address) {
        return new IllegalArgumentException("Data must exactly match destination: "
                + address.rows() + " row(s) x " + address.columns() + " column(s)");
    }

    /** A named single-cell destination. */
    public final class CellInput {
        private final String name;

        private CellInput(String name, Layout.Field ignored) {
            this.name = name;
        }

        /** Buffers a value; a later call for the same name replaces it. */
        public ExcelSheet add(Object value) {
            book.ensureOpen();
            validateValue(value);
            values.put(name, value);
            return ExcelSheet.this;
        }
    }

    /** A named one-row destination. */
    public final class RowInput {
        private final String name;
        private final Layout.Field config;

        private RowInput(String name, Layout.Field config) {
            this.name = name;
            this.config = config;
        }

        /** Buffers one value per configured column. */
        public ExcelSheet add(Object... rowValues) {
            book.ensureOpen();
            int width = config.address().columns();
            if (rowValues == null || rowValues.length != width) {
                throw new IllegalArgumentException(name + ": row requires " + width + " values");
            }
            for (Object value : rowValues) validateValue(value);
            values.put(name, List.of(new ArrayList<>(Arrays.asList(rowValues))));
            return ExcelSheet.this;
        }
    }

    /** A named one-column destination. */
    public final class ColInput {
        private final String name;
        private final Layout.Field config;

        private ColInput(String name, Layout.Field config) {
            this.name = name;
            this.config = config;
        }

        /** Buffers one value per configured row. */
        public ExcelSheet add(Object... columnValues) {
            book.ensureOpen();
            int height = config.address().rows();
            if (columnValues == null || columnValues.length != height) {
                throw new IllegalArgumentException(name + ": column requires " + height + " values");
            }
            List<List<Object>> matrix = new ArrayList<>(height);
            for (Object value : columnValues) {
                validateValue(value);
                matrix.add(new ArrayList<>(Collections.singletonList(value)));
            }
            values.put(name, matrix);
            return ExcelSheet.this;
        }
    }

}
