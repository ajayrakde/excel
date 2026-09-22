package io.github.ajayrakde.excel;

import java.io.IOException;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/** Immutable mapping of business keys to addresses, independent of any sheet. */
public final class ExcelMapper {
    private final Map<String, String> addresses;

    private ExcelMapper(Map<String, String> mappings) {
        if (mappings == null) throw new IllegalArgumentException("Mappings must not be null");
        Map<String, String> checked = new HashMap<>();
        mappings.forEach((key, address) -> {
            if (key == null || key.isBlank()) throw new IllegalArgumentException("Mapping key must not be blank");
            Address.parse(address);
            checked.put(key, address.trim());
        });
        addresses = Map.copyOf(checked);
    }

    public static ExcelMapper of(Map<String, String> mappings) {
        return new ExcelMapper(mappings);
    }

    /** Loads UTF-8 Java properties; duplicate keys are rejected. */
    public static ExcelMapper load(Path path) throws IOException {
        Properties properties = new Properties() {
            @Override public synchronized Object put(Object key, Object value) {
                if (containsKey(key)) throw new IllegalArgumentException("Duplicate mapping key: " + key);
                return super.put(key, value);
            }
        };
        try (Reader reader = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            properties.load(reader);
        }
        Map<String, String> mappings = new HashMap<>();
        properties.stringPropertyNames().forEach(key -> mappings.put(key, properties.getProperty(key)));
        return of(mappings);
    }

    public String resolve(String key) {
        if (key == null || !addresses.containsKey(key)) {
            throw new IllegalArgumentException("Unknown mapping key: " + key);
        }
        return addresses.get(key);
    }
}
