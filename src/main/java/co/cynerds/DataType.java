package co.cynerds;

public enum DataType {
    HOME, NHOME;

    public static DataType fromString(String typeString) {
        return DataType.valueOf(typeString.toUpperCase());
    }
}
