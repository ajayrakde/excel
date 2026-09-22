package io.github.ajayrakde.excel;

import java.util.Locale;
import java.util.regex.Pattern;

/** Internal, zero-based, inclusive XLSX rectangle. */
record Address(int firstRow, int firstColumn, int lastRow, int lastColumn) {
    private static final Pattern CELL = Pattern.compile("\\$?[A-Z]{1,3}\\$?[1-9][0-9]{0,6}");

    static Address parse(String text) {
        if (text == null) throw new IllegalArgumentException("Address must not be null");
        String[] parts = text.trim().toUpperCase(Locale.ROOT).split(":", -1);
        if (parts.length < 1 || parts.length > 2) throw invalid(text);
        int[] start = cell(parts[0], text);
        int[] end = parts.length == 1 ? start : cell(parts[1], text);
        if (end[0] < start[0] || end[1] < start[1]) throw invalid(text);
        return new Address(start[0], start[1], end[0], end[1]);
    }

    private static int[] cell(String part, String original) {
        if (!CELL.matcher(part).matches()) throw invalid(original);
        String clean = part.replace("$", "");
        int column = 0;
        int i = 0;
        while (i < clean.length() && Character.isLetter(clean.charAt(i))) {
            column = column * 26 + clean.charAt(i++) - 'A' + 1;
        }
        int row = Integer.parseInt(clean.substring(i));
        if (row > 1_048_576 || column > 16_384) throw invalid(original);
        return new int[] {row - 1, column - 1};
    }

    private static IllegalArgumentException invalid(String text) {
        return new IllegalArgumentException("Invalid XLSX cell/range: " + text
                + ". Use an address such as B6 or A1:D3 without a sheet name.");
    }

    int rows() { return lastRow - firstRow + 1; }
    int columns() { return lastColumn - firstColumn + 1; }
}
