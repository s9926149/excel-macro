package co.cynerds;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.IntStream;

public class CopyPasteBot {

    private final String sourceDir;

    private String targetFilePath;

    private List<SourceExcelFile> sourceFiles;

    private static final int CYCLE = 16;

    /**
     * 0-based, [2-11] [18-27] [34-43]
     */
    private static final int SOURCE_START_ROW_INDEX = 2;

    public List<Integer> baseSourceRows = IntStream.range(
            SOURCE_START_ROW_INDEX, SOURCE_START_ROW_INDEX + 10
    ).boxed().toList();

    /**
     * 0-based, C-K -> 2-10
     */
    private static final int SOURCE_START_COLUMN_INDEX = 2;

    public List<Integer> sourceColumns = IntStream.range(
            SOURCE_START_COLUMN_INDEX, SOURCE_START_COLUMN_INDEX + 9
    ).boxed().toList();

    private static final CellCopyPolicy COPY_POLICY = new CellCopyPolicy.Builder()
            .cellValue(true)
            .cellStyle(false)
            .cellFormula(false)
            .copyHyperlink(false)
            .mergeHyperlink(false)
            .rowHeight(false)
            .condenseRows(false)
            .mergedRegions(false)
            .build();

    public CopyPasteBot(String sourceDir) {
        this.sourceDir = sourceDir;
        this.sourceFiles = SourceExcelFileListFinder.getSourceExcelFilesAndCheckFileCount(sourceDir);
    }

    public String pasteTo(String targetTemplateFilename) throws IOException {
        this.targetFilePath = getTargetFilePath(targetTemplateFilename);

        try (Workbook wb = WorkbookFactory.create(new File(targetTemplateFilename))) {
            for (SourceExcelFile file : this.sourceFiles) {
                copyAndPasteSingleFile(file, wb);
            }
        }

        return this.targetFilePath;
    }

    public void copyAndPasteSingleFile(SourceExcelFile file, Workbook templateWorkbook) throws IOException {
        try (Workbook wb = WorkbookFactory.create(file.getPath().toFile())) {
            Sheet sheet = wb.getSheet(SourceExcelFile.getSourceSheetName());

            List<CopiedCell> copiedCells = new ArrayList<>();

            // 3 sections need to copy, 16 rows form a cycle
            for (Integer i = 0; i < 3; i++) {
                for (Integer baseSourceRow : baseSourceRows) {
                    Integer rowNum = baseSourceRow + (i * CYCLE);
                    Row row = sheet.getRow(rowNum);

                    for (Integer colNum : sourceColumns) {
                        Cell cell = row.getCell(colNum);
                        copiedCells.add(new CopiedCell(rowNum, colNum, cell, file.getType(), file.getTestDateOrder()));
                    }
                }
            }

            paste(copiedCells, templateWorkbook);
        }
    }

    private void paste(List<CopiedCell> copiedCells, Workbook templateWorkbook) throws IOException {
        for (CopiedCell copiedCell : copiedCells) {
            Sheet sheet = templateWorkbook.getSheet(getTargetSheet(copiedCell));
            pasteCell(sheet, copiedCell);
        }

        try (FileOutputStream out = new FileOutputStream(this.targetFilePath)) {
            templateWorkbook.write(out);
        }
    }

    private static String getTargetSheet(CopiedCell cell) {
        return "Day%d".formatted(cell.testDateOrder());
    }

    private static void pasteCell(Sheet targetSheet, CopiedCell sourceCell) {
        Row row = targetSheet.getRow(getTargetRow(sourceCell));

        if (Objects.nonNull(row)) {
            Cell targetCell = row.getCell(getTargetColumn(sourceCell));
            if (Objects.nonNull(targetCell)) {
                CellUtil.copyCell(sourceCell.cell(), targetCell, COPY_POLICY, null);
            }
        }
    }

    private static int getTargetRow(CopiedCell sourceCell) {
        return sourceCell.sourceRow() + 7; // 3 shift to 10
    }

    private static int getTargetColumn(CopiedCell sourceCell) {
        return switch (sourceCell.type()) {
            case HOME:
                yield sourceCell.sourceCol() - 1; // C shifts to B, K shifts to J
            case NHOME:
                yield sourceCell.sourceCol() + 10; // C shifts to M, K shifts to U
        };
    }

    public static String getTargetFilePath(String targetTemplateFilename) {
        Path templatePath = Paths.get(targetTemplateFilename).toAbsolutePath();

        return Paths.get(templatePath.getParent().toString(), "!" + templatePath.getFileName().toString()).toString();
    }
}
