package com.spring.excel.helper;

import com.spring.excel.model.Tutorial;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelHelper {
    public static String TYPE="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] headers={"Id","Title","Description","Published"};
    static String SHEET="tutorials";

    public static boolean hasExcelFormat(MultipartFile file){
        if(!TYPE.equals(file.getContentType())){
            return false;
        }
        return true;
    }

    public static List<Tutorial> excelToTutorials(InputStream inputStream) {
        try {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(SHEET);
            Iterator<Row> rows = sheet.iterator();
            List<Tutorial> tutorials = new ArrayList<>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();
                //skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }
                Iterator<Cell> cellsInRow = currentRow.iterator();
                Tutorial tutorial = new Tutorial();
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();
                    switch (currentCell.getColumnIndex()) {
                        case 0:
                            tutorial.setId(Double.doubleToLongBits(currentCell.getNumericCellValue()));
                            break;
                        case 1:
                            tutorial.setTitle(currentCell.getStringCellValue());
                            break;
                        case 2:
                            tutorial.setDescription(currentCell.getStringCellValue());
                            break;
                        case 3:
                            tutorial.setPublished(currentCell.getBooleanCellValue());
                            break;

                        default:
                            break;
                    }

                }
                tutorials.add(tutorial);
            }

                workbook.close();

                return tutorials;
            
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
            
        }
        
    }
    
    public static ByteArrayInputStream tutorialToExcel(List<Tutorial> tutorials) throws Exception {
    	ByteArrayOutputStream outputStream=new ByteArrayOutputStream();
    	Workbook workbook=new XSSFWorkbook();
    	try {
    		Sheet sheet=workbook.createSheet(SHEET);
    		Row row=sheet.createRow(0);
    		
    		for( int i=0;i<headers.length;i++) {
    			Cell cell=row.createCell(i);
    			cell.setCellValue(headers[i]);
    			
    		}
    		
    		int rowIndex=1;
    		for(Tutorial eachTutorial:tutorials) {
    			Row dataRow=sheet.createRow(rowIndex);
    			dataRow.createCell(0).setCellValue(eachTutorial.getId());
    			dataRow.createCell(1).setCellValue(eachTutorial.getTitle());
    			dataRow.createCell(2).setCellValue(eachTutorial.getDescription());
    			dataRow.createCell(3).setCellValue(eachTutorial.isPublished());
    		}
    		workbook.write(outputStream);
    		return new ByteArrayInputStream(outputStream.toByteArray());
    		
    	}catch (Exception e) {
    		throw new Exception("Unable to write data to excel file");
		}
    	finally {
			workbook.close();
			outputStream.close();
		}
    	}
}
