package co.cynerds;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class SourceExcelFileListFinder {

    private static final Set<String> FILES_MUST_CONTAIN = Collections.unmodifiableSet(
            Set.of("HOME1", "HOME2", "HOME3", "NHOME1", "NHOME2", "NHOME3")
    );

    private static final String MSG_TEMPLATE = "Found files are not expect as HOME1 to HOME3 and NHOME1 to NHOME3, found: \n";

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
        Set<String> fileKeys = files.stream()
                .map(f -> f.getType().toString() + f.getTestDateOrder())
                .collect(Collectors.toSet());

        return FILES_MUST_CONTAIN.equals(fileKeys);
    }

}
