package co.cynerds;

import java.nio.file.Path;
import java.util.Objects;
import java.util.StringJoiner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SourceExcelFile {

    private final Path path;

    private final DataType type;

    private final Integer testDateOrder;

    private final Integer loopCount;

    /**
     * 0-based, B-J -> 1-9
     */
    private static final int TARGET_HOME_START_COLUMN_INDEX = 1;

    /**
     * 0-based, M-U -> 12-20
     */
    private static final int TARGET_NHOME_START_COLUMN_INDEX = 12;

    /**
     * 0-based, [9-18] [25-34] [41-50]
     */
    private static final int TARGET_START_ROW_INDEX = 9;

    public static final Pattern SOURCE_FILE_NAME_PATTERN =
            Pattern.compile("^\\[Result\\](?<loopCount>\\d*)_Loop_10.*\\[.*_(?<type>HOME|NHOME)(?<testDateOrder>\\d)?\\]_\\d*\\.xls$");

    private static final String SOURCE_SHEET_NAME = "DataTable";


    public SourceExcelFile(Path path, DataType type, Integer testDateOrder, Integer loopCount) {
        this.path = path;
        this.type = type;
        this.testDateOrder = testDateOrder;
        this.loopCount = loopCount;
    }

    public static SourceExcelFile fromPath(Path path) {
        checkIfFileExists(path);

        Matcher matcher = SOURCE_FILE_NAME_PATTERN.matcher(path.getFileName().toString());

        if (matcher.matches()) {
            try {
                String testDateOrder = matcher.group("testDateOrder");
                String loopCount = matcher.group("loopCount");
                return new SourceExcelFile(
                        path,
                        DataType.valueOf(matcher.group("type")),
                        Objects.nonNull(testDateOrder) ? Integer.valueOf(testDateOrder) : null,
                        Objects.nonNull(loopCount) ? Integer.valueOf(loopCount) : null
                );
            } catch (IllegalArgumentException e) {
                System.out.printf(
                        "Warning: <%s> is not a correct source file name, DataType: <%s>, TestDateOrder: <%s>, LoopCount: <%s> %n",
                        path.getFileName().toString(),
                        matcher.group("type"),
                        matcher.group("testDateOrder"),
                        matcher.group("loopCount")

                );
                return null;
            }
        } else {
            return null;
        }

    }

    private static void checkIfFileExists(Path path) {
        if (!path.toFile().exists()) {
            throw new IllegalArgumentException("No such file at %s".formatted(path.toAbsolutePath().toString()));
        }
    }

    public Path getPath() {
        return path;
    }

    public DataType getType() {
        return type;
    }

    public Integer getTestDateOrder() {
        return testDateOrder;
    }

    public Integer getLoopCount() {
        return loopCount;
    }

    public int getTargetStartColumnIndex() {
        if (DataType.HOME.equals(this.type)) {
            return TARGET_HOME_START_COLUMN_INDEX;
        } else {
            return TARGET_NHOME_START_COLUMN_INDEX;
        }
    }

    public static String getSourceSheetName() {
        return SOURCE_SHEET_NAME;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", SourceExcelFile.class.getSimpleName() + ": {", "}")
                .add("\"path\": " + path)
                .add("\"type\": " + type)
                .add("\"testDateOrder\": " + testDateOrder)
                .add("\"loopCount\": " + loopCount)
                .toString();
    }
}
