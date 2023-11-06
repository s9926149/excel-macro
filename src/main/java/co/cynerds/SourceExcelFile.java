package co.cynerds;

import java.nio.file.Path;
import java.util.StringJoiner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SourceExcelFile {

    private final Path path;

    private final DataType type;

    private final int testDateOrder;

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

    public static final Pattern SOURCE_FILE_NAME_PATTERN = Pattern.compile("^\\[Result\\].*\\[.*_(.*)(\\d)\\]_\\d*\\.xls$");

    private static final String SOURCE_SHEET_NAME = "DataTable";


    public SourceExcelFile(Path path, DataType type, int testDateOrder) {
        this.path = path;
        this.type = type;
        this.testDateOrder = testDateOrder;
    }

    public static SourceExcelFile fromPath(Path path) {
        Matcher matcher = SOURCE_FILE_NAME_PATTERN.matcher(path.getFileName().toString());

        if (matcher.matches()) {
            try {
                return new SourceExcelFile(
                        path,
                        DataType.valueOf(matcher.group(1)),
                        Integer.parseInt(matcher.group(2))
                );
            } catch (IllegalArgumentException e) {
                System.out.printf(
                        "Warning: <%s> is not a correct source file name, DataType: <%s>, TestDateOrder: <%s> %n",
                        path.getFileName().toString(),
                        matcher.group(1),
                        matcher.group(2)
                );
                return null;
            }
        } else {
            return null;
        }

    }

    public Path getPath() {
        return path;
    }

    public DataType getType() {
        return type;
    }

    public int getTestDateOrder() {
        return testDateOrder;
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
                .add("\"path\": \"" + path + "\"")
                .add("\"type\": " + type)
                .add("\"testDateOrder\": " + testDateOrder)
                .toString();
    }
}
