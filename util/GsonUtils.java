package com.antaiib.custom.utils;

import com.google.gson.*;
import com.google.gson.reflect.TypeToken;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;

import java.lang.reflect.Type;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * gson工具类
 *
 * @author Administrator
 * @since 2024/02/27 09:27
 */
@UtilityClass
public class GsonUtils {

    private static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final DateTimeFormatter TIME_FORMATTER = DateTimeFormatter.ofPattern("HH:mm:ss");

    private static final JsonSerializer<LocalDateTime> DATE_TIME_SERIALIZER
            = (obj, type, ctx) -> new JsonPrimitive(DATE_TIME_FORMATTER.format(obj));
    private static final JsonSerializer<LocalDate> DATE_SERIALIZER
            = (obj, type, ctx) -> new JsonPrimitive(DATE_FORMATTER.format(obj));
    private static final JsonSerializer<LocalTime> TIME_SERIALIZER
            = (obj, type, ctx) -> new JsonPrimitive(TIME_FORMATTER.format(obj));

    private static final JsonDeserializer<LocalDateTime> DATE_TIME_DESERIALIZER
            = (json, type, ctx) -> LocalDateTime.parse(json.getAsJsonPrimitive().getAsString(), DATE_TIME_FORMATTER);
    private static final JsonDeserializer<LocalDate> DATE_DESERIALIZER
            = (json, type, ctx) -> LocalDate.parse(json.getAsJsonPrimitive().getAsString(), DATE_FORMATTER);
    private static final JsonDeserializer<LocalTime> TIME_DESERIALIZER
            = (json, type, ctx) -> LocalTime.parse(json.getAsJsonPrimitive().getAsString(), TIME_FORMATTER);

    private static final Gson GSON;
    private static final JsonParser JSON_PARSER = new JsonParser();

    static {
        GsonBuilder builder = new GsonBuilder();
        builder.disableHtmlEscaping();
        builder.enableComplexMapKeySerialization();
        builder.setDateFormat("yyyy-MM-dd HH:mm:ss");
        builder.registerTypeAdapter(LocalDateTime.class, DATE_TIME_SERIALIZER);
        builder.registerTypeAdapter(LocalDate.class, DATE_SERIALIZER);
        builder.registerTypeAdapter(LocalTime.class, TIME_SERIALIZER);
        builder.registerTypeAdapter(LocalDateTime.class, DATE_TIME_DESERIALIZER);
        builder.registerTypeAdapter(LocalDate.class, DATE_DESERIALIZER);
        builder.registerTypeAdapter(LocalTime.class, TIME_DESERIALIZER);
        GSON = builder.create();
    }

    /**
     * 返回一个GSON实例
     *
     * @return {@link Gson}
     */
    public static Gson gson() {
        return GSON;
    }

    public static Type makeJavaType(Type rawType, Type... typeArguments) {
        return TypeToken.getParameterized(rawType, typeArguments).getType();
    }

    public static String toString(Object value) {
        if (Objects.isNull(value)) {
            return null;
        }
        if (value instanceof String) {
            return (String) value;
        }
        return toJsonString(value);
    }

    public static String toJsonString(Object value) {
        return GSON.toJson(value);
    }

    public static String toPrettyString(Object value) {
        return GSON.newBuilder().setPrettyPrinting().create().toJson(value);
    }

    public static JsonElement fromJavaObject(Object value) {
        JsonElement result = null;
        if (Objects.nonNull(value) && (value instanceof String)) {
            result = parseObject((String) value);
        } else {
            result = GSON.toJsonTree(value);
        }
        return result;
    }

    @SneakyThrows
    public static JsonElement parseObject(String content) {
        return JSON_PARSER.parse(content);
    }

    public static JsonElement getJsonElement(JsonObject node, String name) {
        return node.get(name);
    }

    public static JsonElement getJsonElement(JsonArray node, int index) {
        return node.get(index);
    }

    @SneakyThrows
    public static <T> T toJavaObject(JsonElement node, Class<T> clazz) {
        return GSON.fromJson(node, clazz);
    }

    @SneakyThrows
    public static <T> T toJavaObject(JsonElement node, Type type) {
        return GSON.fromJson(node, type);
    }

    public static <T> T toJavaObject(JsonElement node, TypeToken<?> typeToken) {
        return toJavaObject(node, typeToken.getType());
    }

    public static <E> List<E> toJavaList(JsonElement node, Class<E> clazz) {
        return toJavaObject(node, makeJavaType(List.class, clazz));
    }

    public static List<Object> toJavaList(JsonElement node) {
        return toJavaObject(node, new TypeToken<List<Object>>(){}.getType());
    }

    public static <V> Map<String, V> toJavaMap(JsonElement node, Class<V> clazz) {
        return toJavaObject(node, makeJavaType(Map.class, String.class, clazz));
    }

    public static Map<String, Object> toJavaMap(JsonElement node) {
        return toJavaObject(node, new TypeToken<Map<String, Object>>(){}.getType());
    }

    @SneakyThrows
    public static <T> T toJavaObject(String content, Class<T> clazz) {
        return GSON.fromJson(content, clazz);
    }

    @SneakyThrows
    public static <T> T toJavaObject(String content, Type type) {
        return GSON.fromJson(content, type);
    }

    public static <T> T toJavaObject(String content, TypeToken<?> typeToken) {
        return toJavaObject(content, typeToken.getType());
    }

    public static <E> List<E> toJavaList(String content, Class<E> clazz) {
        return toJavaObject(content, makeJavaType(List.class, clazz));
    }

    public static List<Object> toJavaList(String content) {
        return toJavaObject(content, new TypeToken<List<Object>>(){}.getType());
    }

    public static <V> Map<String, V> toJavaMap(String content, Class<V> clazz) {
        return toJavaObject(content, makeJavaType(Map.class, String.class, clazz));
    }

    /**
     * JSON字符串转MAP
     *
     * @param content JSON字符串
     * @return {@link Map}<{@link String}, {@link Object}>
     */
    public static Map<String, Object> toJavaMap(String content) {
        return toJavaObject(content, new TypeToken<Map<String, Object>>(){}.getType());
    }

}
