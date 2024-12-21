package com.kvm.Service;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelDataProcessorService {
    static String category;
    static String sw;

    public void processExcelData(File inputFile, File outputZipFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet
            if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                throw new IllegalArgumentException("The sheet is empty or missing.");
            }

            // Processing header row and data rows
            Row headerRow = sheet.getRow(3);
            int assigneeIndex = -1, dateIndex = -1, executedEAIndex = -1, executedRepeatCountIndex = -1, mdIndex = -1, categoryIndex = -1, swIndex = -1;

            // Extracting column indices based on headers
            for (Cell cell : headerRow) {
                String header = cell.getStringCellValue();
                if (header.trim().equalsIgnoreCase("Assignee")) assigneeIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("Due Date")) dateIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("Executed (EA)")) executedEAIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("Executed (Repeat Count)")) executedRepeatCountIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("MD")) mdIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("Category")) categoryIndex = cell.getColumnIndex();
                if (header.trim().equalsIgnoreCase("Category (SW)")) swIndex = cell.getColumnIndex();
            }

            Map<String, Map<String, Map<String, Double>>> results = new HashMap<>();

            // Iterating through rows and processing data
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String assignee = getCellValueAsString(row.getCell(assigneeIndex));
                String date = getCellValueAsString(row.getCell(dateIndex));
                Double executedEA = getCellValueAsDouble(row.getCell(executedEAIndex));
                Double executedRepeatCount = getCellValueAsDouble(row.getCell(executedRepeatCountIndex));
                Double md = getCellValueAsDouble(row.getCell(mdIndex));
                category = getCellValueAsString(row.getCell(categoryIndex));
                sw = getCellValueAsString(row.getCell(swIndex));

                if (assignee == null || date == null || executedEA == null || executedRepeatCount == null || md == null || md == 0) {
                    continue;
                }

                double result = (executedEA * executedRepeatCount) / md;

                results.computeIfAbsent(assignee, k -> new HashMap<>())
                        .computeIfAbsent(date, k -> new HashMap<>())
                        .put(date, result);
            }

            // Writing results to a zip file
            writeResultsToZip(outputZipFile, results);

        }
    }

    private void writeResultsToZip(File zipFile, Map<String, Map<String, Map<String, Double>>> results) throws IOException {
        final double BENCHMARK = 115.0;

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {

            for (String assignee : results.keySet()) {
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                try (Workbook workbook = new XSSFWorkbook()) {
                    Sheet sheet = workbook.createSheet("Results");

                    Row headerRow = sheet.createRow(3);
                    headerRow.createCell(0).setCellValue("Date");
                    headerRow.createCell(1).setCellValue("Category");
                    headerRow.createCell(2).setCellValue("Category (SW)");
                    headerRow.createCell(3).setCellValue("Test Cases Executed");
                    headerRow.createCell(4).setCellValue("Achieved (%)");
                    headerRow.createCell(5).setCellValue("Not Achieved (%)");

                    int rowNum = 4;
                    Map<String, Map<String, Double>> assigneeData = results.get(assignee);
                    for (Map.Entry<String, Map<String, Double>> dateEntry : assigneeData.entrySet()) {
                        String date = dateEntry.getKey();
                        double result = dateEntry.getValue().get(date);

                        double percentageAbove = 0.0;
                        double percentageBelow = 0.0;

                        if (result > BENCHMARK) {
                            percentageAbove = ((result - BENCHMARK) / BENCHMARK) * 100;
                        } else if (result < BENCHMARK) {
                            percentageBelow = ((BENCHMARK - result) / BENCHMARK) * 100;
                        }

                        CellStyle boldStyle = workbook.createCellStyle();
                        Font boldFont = workbook.createFont();
                        boldFont.setBold(true);
                        boldStyle.setFont(boldFont);

                        // Write data to the row
                        Row row1 = sheet.createRow(0);
                        Cell cell1 = row1.createCell(0);
                        cell1.setCellValue("Assignee: ");
                        cell1.setCellStyle(boldStyle);

                        Cell cell2 = row1.createCell(1);
                        cell2.setCellValue(assignee);
                        cell2.setCellStyle(boldStyle);

                        Row row = sheet.createRow(rowNum++);
                        row.createCell(0).setCellValue(date);
                        row.createCell(1).setCellValue(category);
                        row.createCell(2).setCellValue(sw);
                        row.createCell(3).setCellValue(result);
                        row.createCell(4).setCellValue(percentageAbove);
                        row.createCell(5).setCellValue(percentageBelow);
                    }

                    workbook.write(bos);
                }

                ZipEntry zipEntry = new ZipEntry(assignee + ".xlsx");
                zos.putNextEntry(zipEntry);
                zos.write(bos.toByteArray());
                zos.closeEntry();
            }

        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return null;
        }
    }

    private static Double getCellValueAsDouble(Cell cell) {
        if (cell == null) return null;
        return cell.getCellType() == CellType.NUMERIC ? cell.getNumericCellValue() : null;
    }
}

