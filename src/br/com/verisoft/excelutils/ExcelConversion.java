package br.com.verisoft.excelutils;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import au.com.bytecode.opencsv.CSVReader;

public class ExcelConversion {

    public static void main(String[] args) throws IOException {
        long start = System.currentTimeMillis();
        if (args.length < 3) {
            System.out.println("Usage: java -jar csv2xls.jar [separator] [src-file].csv [dest-file].xlsx [lines (optional, default 100)]");
        } else {
            try {
                int nrLinesInMemory = args.length > 3 ? Integer.parseInt(args[3]) : 100;
                System.out.println("SRC-File: " + args[1] + " | Separator: '" + args[0] + "' | DEST-File: " + args[2] + " | Max Lines in Memory: " + nrLinesInMemory);
                SXSSFWorkbook workbook = new SXSSFWorkbook(nrLinesInMemory);
                CreationHelper creationHelper = workbook.getCreationHelper();
                Sheet sheet = workbook.createSheet();
                CSVReader csvReader = new CSVReader(new FileReader(args[1]), args[0].charAt(0));
                String[] csvLineContent;
                int rowNumber = 0;
                while ((csvLineContent = csvReader.readNext()) != null) {
                    Row row = sheet.createRow(rowNumber++);
                    if (rowNumber == 0) {
                        for (int columnNumber = 0; columnNumber < sheet.getRow(0).getPhysicalNumberOfCells(); columnNumber++) {
                            sheet.autoSizeColumn(columnNumber);
                        }
                    }
                    for (int cellNumber = 0; cellNumber < csvLineContent.length; cellNumber++) {
                        Cell cell = row.createCell(cellNumber);
                        cell.setCellValue(creationHelper.createRichTextString(csvLineContent[cellNumber]));
                    }
                }
                for (int columnNumber = 0; columnNumber < sheet.getRow(sheet.getPhysicalNumberOfRows() - 1).getPhysicalNumberOfCells(); columnNumber++) {
                    int columnWidth = sheet.getColumnWidth(columnNumber);
                    sheet.autoSizeColumn(columnNumber);
                    if (sheet.getColumnWidth(columnNumber) < columnWidth) {
                        sheet.setColumnWidth(columnNumber, columnWidth);
                    }
                }
                csvReader.close();
                FileOutputStream fileOutputStream = new FileOutputStream(args[2]);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                workbook.dispose();
            } catch (Exception exception) {
                exception.printStackTrace();
            }
        }
        System.out.println("Conversion finished! Processing time: " + (System.currentTimeMillis() - start) + "ms.");
    }

}
