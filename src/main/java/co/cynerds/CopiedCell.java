package co.cynerds;

import org.apache.poi.ss.usermodel.Cell;

public record CopiedCell(int sourceRow, int sourceCol, Cell cell, DataType type, int testDateOrder) {}
