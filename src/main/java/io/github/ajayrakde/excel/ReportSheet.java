package io.github.ajayrakde.excel;

import java.util.LinkedHashMap;
import java.util.Map;

/** Named cells and tables on one configured worksheet. */
public final class ReportSheet {
    private final ExcelBook book;
    private final ExcelSheet sheet;
    private final Map<String, ReportLayout.Field> fields;
    private final Map<String, Object> values = new LinkedHashMap<>();
    private final Map<String, ReportTable> tables = new LinkedHashMap<>();

    ReportSheet(ExcelBook book, ExcelSheet sheet, Map<String, ReportLayout.Field> fields) {
        this.book = book;
        this.sheet = sheet;
        this.fields = fields;
    }

    /** Buffers a scalar; repeated calls to the same key replace the pending value. */
    public ReportSheet set(String name, Object value) {
        field(name, "cell");
        ExcelSheet.validateValue(value);
        values.put(name, value);
        return this;
    }

    /** Returns the same table builder on repeated calls, so rows continue appending. */
    public ReportTable table(String name) {
        var config = field(name, "table");
        return tables.computeIfAbsent(name, ignored -> new ReportTable(book, name, config));
    }

    private ReportLayout.Field field(String name, String expected) {
        book.ensureOpen();
        var field = name == null ? null : fields.get(name);
        if (field == null) throw new IllegalArgumentException("Unknown field: " + name);
        if (!field.type().equals(expected)) {
            throw new IllegalArgumentException(name + " is a " + field.type() + ", not a " + expected);
        }
        return field;
    }

    void validate() {
        values.forEach((name, value) -> sheet.validate(fields.get(name).range(), value));
        tables.values().forEach(table -> table.validate(sheet));
    }

    void apply() {
        values.forEach((name, value) -> sheet.write(fields.get(name).range(), value));
        tables.values().forEach(table -> table.apply(sheet));
    }
}
