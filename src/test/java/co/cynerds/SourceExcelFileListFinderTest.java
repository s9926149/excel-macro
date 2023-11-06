package co.cynerds;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SourceExcelFileListFinderTest {

    @Test
    public void testNotExactFileCount() {
        IllegalArgumentException illegalArgumentException = assertThrows(IllegalArgumentException.class, () -> {
            SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample/file-count-check");
        });
    }

    @Test
    void testExactFileCount() {
        assertEquals(6, SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount("./sample").size());
    }

}