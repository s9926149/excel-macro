package co.cynerds;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SourceExcelFileListFinderTest {

    @Test
    public void testNotExactFileCount() {
        IllegalArgumentException illegalArgumentException = assertThrows(IllegalArgumentException.class, () -> {
            SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/file-count-check/invalid-incomplete-files");
        });
    }

    @Test
    public void testTooManyFileCount() {
        IllegalArgumentException illegalArgumentException = assertThrows(IllegalArgumentException.class, () -> {
            SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/file-count-check/invalid-too-many-files");
        });
    }

    @Test
    void testExactFileCountAll3LoopsVersion() {
        assertEquals(6, SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/3-loops-version").size());
    }

    @Test
    void testExactFileCountAll10LoopsFrom0Version() {
        assertEquals(
                2,
                SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/10-loops-start-from-0-version").size()
        );
    }

    @Test
    void testExactFileCountAll10LoopsFrom1Version() {
        assertEquals(
                2,
                SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/10-loops-start-from-1-version").size()
        );
    }

    @Test
    void testExactFileCount10LoopsHomeFrom0Version3LoopsNhome() {
        assertEquals(
                4,
                SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/file-count-check/valid-10-loops-home-from-0-ver-3-loops-nhome").size()
        );
    }

}