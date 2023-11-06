package co.cynerds;

import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.*;

class SourceExcelFileTest {

    @Test
    public void testFromPathWithCorrectPath() {
        String filePath = "./sample/[Result]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME1]_20231002175041.xls";
        Path path = Paths.get(filePath);

        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertEquals(path.toString(), file.getPath().toString());
        assertEquals(DataType.HOME, file.getType());
        assertEquals(1, file.getTestDateOrder());
    }

    @Test
    public void testFromPathWithNotMatchPath() {
        String filePath =
                "./sample/[Resultss]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOME1]_20231002175041.xls";
        Path path = Paths.get(filePath);
        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertNull(file);
    }

    @Test
    public void testFromPathWithIncorrectDataType() {
        String filePath =
                "./sample/[Result]3_Loop_10_Cycle_All_result_BL4000_SN[NPT22-00XX_HOMEZ1]_20231002175041.xls";
        Path path = Paths.get(filePath);
        SourceExcelFile file = SourceExcelFile.fromPath(path);

        assertNull(file);
    }

}