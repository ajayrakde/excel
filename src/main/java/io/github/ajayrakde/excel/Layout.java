package io.github.ajayrakde.excel;

import java.io.IOException;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.HashSet;
import org.yaml.snakeyaml.LoaderOptions;
import org.yaml.snakeyaml.Yaml;
import org.yaml.snakeyaml.constructor.SafeConstructor;

/** Parses and validates the internal YAML layout model. */
final class Layout {
    record Field(String type, Address address, boolean hasHeader, Integer limit, List<String> headers) {}

    static Map<String, Map<String, Field>> load(Path path) throws IOException {
        LoaderOptions options = new LoaderOptions();
        options.setAllowDuplicateKeys(false);
        options.setMaxAliasesForCollections(0);
        Object document;
        try (Reader reader = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            document = new Yaml(new SafeConstructor(options)).load(reader);
        } catch (org.yaml.snakeyaml.error.YAMLException error) {
            throw new IllegalArgumentException("Invalid layout YAML: " + error.getMessage(), error);
        }
        Map<?, ?> root = mapping(document, "layout");
        keys(root, Set.of("sheets"), "layout");
        Map<?, ?> sheets = mapping(root.get("sheets"), "sheets");
        Map<String, Map<String, Field>> result = new LinkedHashMap<>();
        for (var entry : sheets.entrySet()) {
            String sheet = text(entry.getKey(), "sheet name");
            Map<?, ?> fields = mapping(entry.getValue(), "sheet " + sheet);
            Map<String, Field> parsed = new LinkedHashMap<>();
            for (var field : fields.entrySet()) {
                String name = text(field.getKey(), "field name");
                String context = sheet + "." + name;
                Map<?, ?> config = mapping(field.getValue(), context);
                String type = text(config.get("type"), context + ".type");
                if (!Set.of("cell", "table").contains(type)) {
                    throw new IllegalArgumentException(context + ": type must be cell or table");
                }
                keys(config, type.equals("table")
                        ? Set.of("type", "range", "hasHeader", "limit", "headers")
                        : Set.of("type", "range"), context);
                String range = text(config.get("range"), context + ".range");
                Address address;
                try { address = Address.parse(range); }
                catch (IllegalArgumentException error) {
                    throw new IllegalArgumentException(context + ": " + error.getMessage(), error);
                }
                if (address.rows() != 1 || (type.equals("cell") && address.columns() != 1)) {
                    throw new IllegalArgumentException(context + ": range must identify "
                            + (type.equals("cell") ? "one cell" : "only the first row, such as A2:C2"));
                }
                boolean hasHeader = false;
                Integer limit = null;
                List<String> headers = List.of();
                if (type.equals("table")) {
                    if (!(config.get("hasHeader") instanceof Boolean header)) {
                        throw new IllegalArgumentException(context + ": hasHeader must be true or false");
                    }
                    hasHeader = header;
                    if (config.containsKey("limit")) {
                        if (!(config.get("limit") instanceof Integer count) || count < 1) {
                            throw new IllegalArgumentException(context + ": limit must be a positive integer");
                        }
                        limit = count;
                    }
                    if (config.containsKey("headers")) {
                        if (!(config.get("headers") instanceof List<?> configured)
                                || configured.size() != address.columns()) {
                            throw new IllegalArgumentException(context + ": headers must contain exactly "
                                    + address.columns() + " names");
                        }
                        var parsedHeaders = new java.util.ArrayList<String>();
                        var unique = new HashSet<String>();
                        for (Object item : configured) {
                            String headerName = text(item, context + ".headers");
                            if (!unique.add(headerName)) {
                                throw new IllegalArgumentException(context + ": duplicate header " + headerName);
                            }
                            parsedHeaders.add(headerName);
                        }
                        headers = List.copyOf(parsedHeaders);
                    }
                    long firstDataRow = address.firstRow() + (hasHeader ? 1L : 0L);
                    if (firstDataRow >= 1_048_576 || (limit != null && firstDataRow + limit > 1_048_576)) {
                        throw new IllegalArgumentException(context + ": table exceeds XLSX row limit");
                    }
                }
                parsed.put(name, new Field(type, address, hasHeader, limit, headers));
            }
            result.put(sheet, Map.copyOf(parsed));
        }
        return Map.copyOf(result);
    }

    private static Map<?, ?> mapping(Object value, String context) {
        if (!(value instanceof Map<?, ?> map) || map.isEmpty()) {
            throw new IllegalArgumentException(context + " must be a non-empty mapping");
        }
        return map;
    }

    private static String text(Object value, String context) {
        if (!(value instanceof String text) || text.isBlank()) {
            throw new IllegalArgumentException(context + " must be a non-blank string");
        }
        return text;
    }

    private static void keys(Map<?, ?> map, Set<String> allowed, String context) {
        for (Object key : map.keySet()) {
            if (!allowed.contains(key)) throw new IllegalArgumentException(context + ": unknown property " + key);
        }
    }
}
