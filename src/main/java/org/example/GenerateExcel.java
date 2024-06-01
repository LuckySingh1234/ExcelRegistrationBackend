package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class GenerateExcel {
    public String addUserToExcel(User user) {
        String[] headers = {"First Name", "Last Name", "Phone", "Mobile"};
        String filePath = "F:/RegisteredUsers.xlsx";
        File directory = new File(filePath.substring(0, filePath.lastIndexOf("/")));
        if (!directory.exists()) {
            directory.mkdirs();
        }
        String sheetName = "Users";
        try {
            Workbook workbook;
            if (Files.exists(Paths.get(filePath))) {
                // If the file exists, open it
                FileInputStream fis = new FileInputStream(filePath);
                workbook = WorkbookFactory.create(fis);
            } else {
                // If the file doesn't exist, create a new workbook
                workbook = new XSSFWorkbook();
            }
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                // Create header row
                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                }
            }
            // Create data rows
            int rowNum = sheet.getPhysicalNumberOfRows();
            int colNum = 0;
            Row row = sheet.createRow(rowNum);
            Cell cell = row.createCell(colNum);
            cell.setCellValue(user.getFirstName());
            colNum++;
            cell = row.createCell(colNum);
            cell.setCellValue(user.getLastName());
            colNum++;
            cell = row.createCell(colNum);
            cell.setCellValue(user.getEmail());
            colNum++;
            cell = row.createCell(colNum);
            cell.setCellValue(user.getPhone());
            // Write the Excel file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
                System.out.println("Excel file created successfully at the below location");
                System.out.println(Paths.get(filePath).toAbsolutePath());
            } catch (IOException e) {
                System.out.println("Error creating Excel file: " + e.getMessage());
            }
        } catch (IOException e) {
            System.out.println("Error creating Excel file: " + e.getMessage());
        }
        return filePath;
    }
}
