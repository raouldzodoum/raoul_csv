package PdfToExcel.Project1;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelHeaderAndBodyExample {
	
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Define header values
        String[] headers = {"Header1", "Header2", "Header3", "Header4"};

        // Create a header row
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Define body data
        Object[][] bodyData = {
                {"Data1", "Data2", "Data3", "Data4"},
                {"Data5", "Data6", "Data7", "Data8"},
                // Add more rows as needed
        };

        // Add body data to the sheet
        int rowNum = 1;
        for (Object[] rowData : bodyData) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < rowData.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(String.valueOf(rowData[i]));
            }
        }

        try {
            // Write the workbook to a file
            FileOutputStream fileOut = new FileOutputStream("output.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Excel file with header and body elements created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
