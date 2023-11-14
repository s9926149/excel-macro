package co.cynerds;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class SourceExcelFileListFinder {

    private static final String MSG_TEMPLATE = "Found files are not expected, found: \n";

    public static List<SourceExcelFile> getSourceExcelFilesAndCheckFileCount(String dir) {
        List<SourceExcelFile> files = getSourceExcelFiles(dir);

        if (isExactRequiredFiles(files)) {
            return files;
        } else {
            List<String> filenames = files.stream()
                    .map(SourceExcelFile::getPath)
                    .map(Path::getFileName)
                    .map(Path::toString)
                    .toList();
            String message = MSG_TEMPLATE + String.join("\n", filenames);
            System.err.println(message);
            throw new IllegalArgumentException();
        }
    }

    public static List<SourceExcelFile> getSourceExcelFiles(String dir) {
        List<Path> filenameList = getFileListUnderDir(dir);

        return filenameList.stream()
                .map(SourceExcelFile::fromPath)
                .filter(Objects::nonNull)
                .collect(Collectors.toList());
    }

    public static List<Path> getFileListUnderDir(String dir) {
        Path path = Paths.get(dir);

        List<Path> filenameList;

        try (Stream<Path> stream = Files.list(path)) {
            filenameList =  stream.filter(f -> !Files.isDirectory(f))
                    .collect(Collectors.toList());
        } catch (IOException e) {
            throw new RuntimeException(
                    "Can't retrieve files from directory: <%s>, please check if it's entered correctly".formatted(dir),
                    e
            );
        }

        return filenameList;
    }

    public static boolean isExactRequiredFiles(List<SourceExcelFile> files) {
        return Arrays.stream(DataType.values())
                .map(t -> check(files, t))
                .reduce(true, (b1, b2) -> b1 && b2);
    }


    private static boolean check(List<SourceExcelFile> files, DataType type) {
        List<String> filesKeysOfType = files.stream()
                .filter(f -> type.equals(f.getType()))
                .map(f -> {
                    String testDateOrder = Objects.nonNull(f.getTestDateOrder()) ? f.getTestDateOrder().toString() : "";
                    return f.getType().toString() + testDateOrder;
                })
                .toList();

        Set<String> FILES_MUST_CONTAIN_3_LOOPS = IntStream.rangeClosed(1, 3)
                .mapToObj(num -> type.name() + num)
                .collect(Collectors.toSet());


        if (
            FILES_MUST_CONTAIN_3_LOOPS.size() == filesKeysOfType.size() &&
            FILES_MUST_CONTAIN_3_LOOPS.containsAll(filesKeysOfType))
        {
            return true;
        } else if (
            filesKeysOfType.size() == 1 && type.name().equals(filesKeysOfType.get(0))
        ) {
            return true;
        } else {
            return false;
        }
    }

}
