import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Arrays;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public final class ExcelFileMerger {

    private static final String EMPTY_STRING = "";

    public static void merge(Sheet targetSheet, int spaceBetweenSheets, boolean direction, List<Sheet> sheetArgs) {
        if(direction) {
            mergeHorizontally(sheetArgs, targetSheet, spaceBetweenSheets);
        } else {
            mergeVertically(sheetArgs, targetSheet, spaceBetweenSheets);
        }
    }

    public static void merge(Sheet targetSheet, int spaceBetweenSheets, boolean direction, Sheet... sheetArgs) {
        if(direction) {
            mergeHorizontally(Arrays.asList(sheetArgs), targetSheet, spaceBetweenSheets);
        } else {
            mergeVertically(Arrays.asList(sheetArgs), targetSheet, spaceBetweenSheets);
        }
    }

    /**
     * Merge all sheets on the given workbook into a new sheet inside the same workbook.
    **/
    public static void merge(Workbook workBook, int spaceBetweenSheets, boolean direction, String sheetName) {
        if(direction) {
            mergeHorizontally(StreamSupport.stream(Spliterators.spliteratorUnknownSize(workBook.sheetIterator(), Spliterator.ORDERED), Boolean.FALSE).collect(Collectors.toList()), workBook.createSheet(sheetName), spaceBetweenSheets);
        } else {
            mergeVertically(StreamSupport.stream(Spliterators.spliteratorUnknownSize(workBook.sheetIterator(), Spliterator.ORDERED), Boolean.FALSE).collect(Collectors.toList()), workBook.createSheet(sheetName), spaceBetweenSheets);
        }
    }

    private static void mergeHorizontally(List<Sheet> sheetArgs, Sheet targetSheet, int spaceBetweenSheets) {
        AtomicInteger rowCount = new AtomicInteger(targetSheet.getLastRowNum());
        sheetArgs.forEach(sheet -> {
            sheet.rowIterator().forEachRemaining(row -> {
                for (int cellNumber = 0; cellNumber < row.getLastCellNum(); cellNumber++) {
                    setCellValue(targetSheet.createRow(rowCount.getAndIncrement()).createCell(cellNumber), row.getCell(cellNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                }
            });
            addEmptyRowsBetweenSheets(targetSheet, rowCount, spaceBetweenSheets);
        });
    }

    private static void mergeVertically(List<Sheet> sheetArgs, Sheet targetSheet, int spaceBetweenSheets) {
        AtomicInteger workbookColumnMaxValue = new AtomicInteger(0);
        sheetArgs.forEach(sheet -> {
            AtomicInteger sheetColumnMaxValue = new AtomicInteger(0);
            sheet.rowIterator().forEachRemaining(row -> {
                Row newRow = targetSheet.getRow(row.getRowNum());

                if(newRow == null) {
                    newRow = targetSheet.createRow(row.getRowNum());
                }

                for (int cellNumber = 0; cellNumber < row.getLastCellNum(); cellNumber++) {
                    setCellValue(newRow.createCell(cellNumber + workbookColumnMaxValue.get()), row.getCell(cellNumber, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                }

                if(row.getLastCellNum() > sheetColumnMaxValue.get()) {
                    sheetColumnMaxValue.set(row.getLastCellNum());
                }
            });
            workbookColumnMaxValue.set(sheetColumnMaxValue.get() + spaceBetweenSheets);
        });
    }

    private static void addEmptyRowsBetweenSheets(Sheet targetSheet, AtomicInteger lastRowTargetSheet, int numberOfLines) {
        while (numberOfLines != 0) {
            targetSheet.createRow(lastRowTargetSheet.getAndIncrement());
            numberOfLines--;
        }
    }

    private static void setCellValue(Cell newCell, Cell oldCell) {
        if(oldCell == null) {
            return;
        }

        switch (oldCell.getCellTypeEnum()) {
            case _NONE:
            case BLANK:
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;

            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;

            case FORMULA:
                newCell.setCellValue(oldCell.getCellFormula());
                break;

            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;

            case ERROR:
                newCell.setCellValue(oldCell.getErrorCellValue());
                break;

            default: newCell.setCellValue(EMPTY_STRING);
        }

        newCell.setCellType(oldCell.getCellTypeEnum());
        newCell.setCellComment(oldCell.getCellComment());
        newCell.setCellStyle(oldCell.getCellStyle());
        newCell.setHyperlink(oldCell.getHyperlink());
    }
}