package co.cynerds;

import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

class SourceExcelFileTest {

    @Test
    public void test3LoopsVerFromPathWithCorrectPath() {
        String filePath = "./sample/3-loops-version/[Result]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME1]_20231002175041.xls";
        Path path = Paths.get(filePath);

        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertEquals(path.toString(), file.getPath().toString());
        assertEquals(DataType.HOME, file.getType());
        assertEquals(1, file.getTestDateOrder());
        assertEquals(3, file.getLoopCount());
    }

    @Test
    public void test10LoopsVerFromPathWithCorrectPath() {
        String filePath = "./sample/10-loops-start-from-0-version/[Result]10_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME]_20231002175041.xls";
        Path path = Paths.get(filePath);

        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertEquals(path.toString(), file.getPath().toString());
        assertEquals(DataType.HOME, file.getType());
        assertNull(file.getTestDateOrder());
        assertEquals(10, file.getLoopCount());
    }

    @Test
    public void testFromPathWithNotMatchPath() {
        String filePath =
                "./sample/invalid-file-names/[Resultss]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME1]_20231002175041.xls";
        Path path = Paths.get(filePath);
        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertNull(file);
    }

    @Test
    public void testFromPathWithIncorrectDataType() {
        String filePath =
                "./sample/invalid-file-names/[Result]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOMEZ1]_20231002175041.xls";
        Path path = Paths.get(filePath);
        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertNull(file);
    }

    @Test
    public void testFromPathWithNoSuchFile() {
        String filePath =
                "./sample/[Results]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME1]_20231002175041.xls";
        Path path = Paths.get(filePath);

        assertThrows(IllegalArgumentException.class, () -> SourceExcelFile.fromPath(path));
    }

}