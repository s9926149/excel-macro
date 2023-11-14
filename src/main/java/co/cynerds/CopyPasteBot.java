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

    private final List<SourceExcelFile> sourceFiles;

    private static final int CYCLE_SIZE = 16;

    /**
     * 0-based, [2-11] [18-27] [34-43]
     */
    private static final int SOURCE_START_ROW_INDEX = 2;

    public final List<Integer> baseSourceRows = IntStream.range(
            SOURCE_START_ROW_INDEX, SOURCE_START_ROW_INDEX + 10
    ).boxed().toList();

    /**
     * 0-based, C-K -> 2-10
     */
    private static final int SOURCE_START_COLUMN_INDEX = 2;

    public final List<Integer> sourceColumns = IntStream.range(
            SOURCE_START_COLUMN_INDEX, SOURCE_START_COLUMN_INDEX + 9
    ).boxed().toList();

    /**
     * 0-based, N -> 13
     */
    public static final int STARTING_SECTOR_DETERMINING_COLUMN = 13;

    /**
     * 0-based, 2
     */
    public static final int STARTING_SECTOR_DETERMINING_ROW = 2;

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

            List<CopiedCell> copiedCells = copyCells(file, sheet);

            paste(copiedCells, templateWorkbook);
        }
    }

    // TODO refactor this pile of...
    public List<CopiedCell> copyCells(SourceExcelFile file, Sheet sheet) {
        Integer loopCount = file.getLoopCount();
        List<CopiedCell> copiedCells = new ArrayList<>();

        if (loopCount == 3) {
            // 3 sections need to copy, 16 rows form a cycle
            for (Integer section = 0; section < 3; section++) {
                for (Integer baseSourceRow : baseSourceRows) {
                    Integer rowNum = baseSourceRow + (section * CYCLE_SIZE);
                    Row row = sheet.getRow(rowNum);

                    for (Integer colNum : sourceColumns) {
                        Cell cell = row.getCell(colNum);
                        copiedCells.add(new CopiedCell(rowNum, colNum, cell, file.getType(), file.getTestDateOrder()));
                    }
                }
            }
        } else if (loopCount == 10) {
            // determine whether starting section is 0 or 1 by specific cell value
            int startingSection = determineStartingSection(sheet);

            // 9 sections need to copy, 16 rows form a cycle
            for (Integer section = startingSection; section < startingSection + 9; section++) {
                for (Integer baseSourceRow : baseSourceRows) {
                    Integer rowNum = baseSourceRow + (section * CYCLE_SIZE);
                    Row row = sheet.getRow(rowNum);

                    int testDateOrder = calculateTestDateOrder(startingSection, section);
                    for (Integer colNum : sourceColumns) {
                        Cell cell = row.getCell(colNum);
                        int shiftedRowNum = shiftRowBackFor10Loops(rowNum, startingSection, testDateOrder);

                        copiedCells.add(new CopiedCell(shiftedRowNum, colNum, cell, file.getType(), testDateOrder));
                    }
                }
            }
        } else {
            throw new IllegalArgumentException(
                    "Only 3 or 10 loop count is accepted, current value: %d".formatted(loopCount)
            );
        }

        return copiedCells;
    }

    private static int determineStartingSection(Sheet sheet) {
        double determiningValue = sheet.getRow(STARTING_SECTOR_DETERMINING_ROW)
                .getCell(STARTING_SECTOR_DETERMINING_COLUMN)
                .getNumericCellValue();
        return determiningValue < 1 ? 0 : 1;
    }

    private static int calculateTestDateOrder(int startingSection, int section) {
        return ((section - startingSection) / 3) + 1;
    }

    private static int shiftRowBackFor10Loops(int rowNum, int startingSection, int testDateOrder) {
        return rowNum - (startingSection * CYCLE_SIZE) - ((testDateOrder - 1) * CYCLE_SIZE * 3);
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
