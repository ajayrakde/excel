package io.github.ajayrakde.excel;

import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/** Writes explicit addresses on one sheet. Not thread-safe. */
public final class ExcelSheet {
    private final ExcelBook book;
    private final Sheet sheet;

    ExcelSheet(ExcelBook book, Sheet sheet) {
        this.book = book;
        this.sheet = sheet;
    }

    /**
     * Writes a scalar, a flat list for a row/column, or a list of rows.
     * Dimensions must match exactly. Null clears a cell. All validation
     * completes before cells are changed; merged overlaps are rejected.
     */
    public ExcelSheet write(String destination, Object data) {
        book.ensureOpen();
        Address address = Address.parse(destination);
        boolean matrix = data instanceof List<?> list && !list.isEmpty()
                && list.get(0) instanceof List<?>;
        validateShape(address, data, matrix);
        for (CellRangeAddress merged : sheet.getMergedRegions()) {
            if (address.firstRow() <= merged.getLastRow() && address.lastRow() >= merged.getFirstRow()
                    && address.firstColumn() <= merged.getLastColumn()
                    && address.lastColumn() >= merged.getFirstColumn()) {
                throw new IllegalArgumentException("Destination " + destination
                        + " overlaps merged cells " + merged.formatAsString());
            }
        }
        for (int r = 0; r < address.rows(); r++) {
            for (int c = 0; c < address.columns(); c++) {
                validateValue(valueAt(address, data, matrix, r, c));
            }
        }
        for (int r = 0; r < address.rows(); r++) {
            Row row = sheet.getRow(address.firstRow() + r);
            if (row == null) row = sheet.createRow(address.firstRow() + r);
            for (int c = 0; c < address.columns(); c++) {
                Cell cell = row.getCell(address.firstColumn() + c);
                if (cell == null) cell = row.createCell(address.firstColumn() + c);
                Object value = valueAt(address, data, matrix, r, c);
                cell.setBlank(); // Also replaces an existing formula; retains style.
                if (value instanceof String text) cell.setCellValue(text);
                else if (value instanceof Boolean bool) cell.setCellValue(bool);
                else if (value instanceof Number number) cell.setCellValue(number.doubleValue());
            }
        }
        return this;
    }

    private static void validateShape(Address address, Object data, boolean matrix) {
        if (data instanceof List<?> values) {
            if (matrix) {
                if (values.size() != address.rows()) throw shape(address);
                for (Object value : values) {
                    if (!(value instanceof List<?> row) || row.size() != address.columns()) throw shape(address);
                }
            } else {
                if (address.rows() != 1 && address.columns() != 1) throw shape(address);
                if (values.size() != Math.max(address.rows(), address.columns())) throw shape(address);
            }
        } else if (address.rows() != 1 || address.columns() != 1) {
            throw shape(address);
        }
    }

    private static IllegalArgumentException shape(Address address) {
        return new IllegalArgumentException("Data must exactly match destination: "
                + address.rows() + " row(s) x " + address.columns() + " column(s)");
    }

    private static Object valueAt(Address address, Object data, boolean matrix, int row, int column) {
        if (!(data instanceof List<?> values)) return data;
        if (matrix) return ((List<?>) values.get(row)).get(column);
        return values.get(address.rows() == 1 ? column : row);
    }

    private static void validateValue(Object value) {
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
}
