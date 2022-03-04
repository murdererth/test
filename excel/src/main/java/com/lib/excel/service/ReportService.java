package com.lib.excel.service;

import com.lib.excel.model.CustomCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Map;

@Service
public class ReportService {

    private StylesGenerator stylesGenerator;

    public ReportService(StylesGenerator stylesGenerator) {
        this.stylesGenerator = stylesGenerator;
    }

    public byte[] generateXlsxReport() {
        Workbook wb = new XSSFWorkbook();

        return generateReport(wb);
    }

    public byte[] generateXlsReport() {
        Workbook wb = new HSSFWorkbook();

        return generateReport(wb);
    }

    private byte[] generateReport(Workbook wb) {
        Map<CustomCellStyle, CellStyle> styles = stylesGenerator.prepareStyles(wb);
        Sheet sheet = wb.createSheet("Example sheet name");

        setColumnsWidth(sheet);

        createHeaderRow(sheet, styles);
        createStringsRow(sheet, styles);
        createDoublesRow(sheet, styles);
        createDatesRow(sheet, styles);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            wb.write(out);
            out.close();
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return out.toByteArray();
    }

    private void setColumnsWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 256 * 20);

        for (int columnIndex = 1; columnIndex <= 5; columnIndex++) {
            sheet.setColumnWidth(columnIndex, 256 * 15);
        }
    }

    private void createHeaderRow(Sheet sheet, Map<CustomCellStyle, CellStyle> styles) {
        Row row = sheet.createRow(0);

        for (int columnNumber = 1; columnNumber <= 5; columnNumber++) {
            Cell cell = row.createCell(columnNumber);

            cell.setCellValue("Column "+columnNumber);
            cell.setCellStyle(styles.get(CustomCellStyle.GREY_CENTERED_BOLD_ARIAL_WITH_BORDER));
        }
    }

    private void createRowLabelCell(Row row, Map<CustomCellStyle, CellStyle> styles, String label) {
        Cell rowLabel = row.createCell(0);
        rowLabel.setCellValue(label);
        rowLabel.setCellStyle(styles.get(CustomCellStyle.RED_BOLD_ARIAL_WITH_BORDER));
    }

    private void createStringsRow(Sheet sheet, Map<CustomCellStyle, CellStyle> styles) {
        Row row = sheet.createRow(1);
        createRowLabelCell(row, styles, "Strings row");

        for (int columnNumber = 1; columnNumber <= 5; columnNumber++) {
            Cell cell = row.createCell(columnNumber);

            cell.setCellValue("String "+columnNumber);
            cell.setCellStyle(styles.get(CustomCellStyle.RIGHT_ALIGNED));
        }
    }

    private void createDoublesRow(Sheet sheet, Map<CustomCellStyle, CellStyle> styles) {
        Row row = sheet.createRow(2);
        createRowLabelCell(row, styles, "Doubles row");

        for (int columnNumber = 1; columnNumber <= 5; columnNumber++) {
            Cell cell = row.createCell(columnNumber);

            cell.setCellValue(Double.parseDouble(columnNumber+".99"));

            cell.setCellStyle(styles.get(CustomCellStyle.RIGHT_ALIGNED));
        }
    }

    private void createDatesRow(Sheet sheet, Map<CustomCellStyle, CellStyle> styles) {
        Row row = sheet.createRow(3);
        createRowLabelCell(row, styles, "Dates row");

        for (int columnNumber = 1; columnNumber <= 5; columnNumber++) {
            Cell cell = row.createCell(columnNumber);

            cell.setCellValue((LocalDate.now()));
            cell.setCellStyle(styles.get(CustomCellStyle.RIGHT_ALIGNED_DATE_FORMAT));
        }
    }
}
