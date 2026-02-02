package com.example.mysheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelMonitor {

    public static void main(String[] args) {
        String filePath = "ScienceFair.xlsx";
        boolean errorsFound = false;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell name = row.getCell(0);
                Cell project = row.getCell(1);
                Cell j1Cell = row.getCell(2);
                Cell j2Cell = row.getCell(3);
                Cell avgCell = row.getCell(4);

                // Rule 1: Name + Project must not be empty
                if (isEmpty(name) || isEmpty(project)) {
                    System.out.println("Row " + (i + 1) + ": Missing name or project");
                    errorsFound = true;
                }

                // Rule 2: Scores must be numeric 0..10 (accept numeric-as-text too)
                Double s1 = getDoubleOrNull(j1Cell);
                Double s2 = getDoubleOrNull(j2Cell);

                if (s1 == null || s1 < 0 || s1 > 10) {
                    System.out.println("Row " + (i + 1) + ": Judge 1 score invalid (0..10)");
                    errorsFound = true;
                }
                if (s2 == null || s2 < 0 || s2 > 10) {
                    System.out.println("Row " + (i + 1) + ": Judge 2 score invalid (0..10)");
                    errorsFound = true;
                }

                // Rule 3: If average exists, validate it (accept numeric-as-text too)
                if (!isEmpty(avgCell) && s1 != null && s2 != null) {
                    double expected = round1((s1 + s2) / 2.0);

                    Double actualAvg = getDoubleOrNull(avgCell);
                    if (actualAvg == null) {
                        System.out.println("Row " + (i + 1) + ": Average is not numeric");
                        errorsFound = true;
                    } else if (Double.compare(round1(actualAvg), expected) != 0) {
                        System.out.println("Row " + (i + 1) + ": Wrong average. Expected "
                                + expected + " but found " + actualAvg);
                        errorsFound = true;
                    }
                }
            }

            if (errorsFound) {
                System.exit(1); // fail CI
            } else {
                System.out.println("All data is correct!");
            }

        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
            System.exit(1);
        }
    }

    static boolean isEmpty(Cell c) {
        if (c == null) return true;
        if (c.getCellType() == CellType.BLANK) return true;

        if (c.getCellType() == CellType.STRING) {
            return c.getStringCellValue() == null || c.getStringCellValue().trim().isEmpty();
        }
        return false;
    }

    static Double getDoubleOrNull(Cell c) {
        if (c == null) return null;

        switch (c.getCellType()) {
            case NUMERIC:
                return c.getNumericCellValue();
            case STRING:
                String s = c.getStringCellValue();
                if (s == null) return null;
                s = s.trim();
                if (s.isEmpty()) return null;
                try {
                    return Double.parseDouble(s);
                } catch (NumberFormatException ex) {
                    return null;
                }
            case FORMULA:
                // If formula returns number, POI type can still be NUMERIC,
                // but safest is to check the cached result:
                try {
                    return c.getNumericCellValue();
                } catch (IllegalStateException ex) {
                    try {
                        String fs = c.getStringCellValue();
                        return (fs == null || fs.trim().isEmpty()) ? null : Double.parseDouble(fs.trim());
                    } catch (Exception ignore) {
                        return null;
                    }
                }
            default:
                return null;
        }
    }

    static double round1(double v) {
        return Math.round(v * 10.0) / 10.0;
    }
}