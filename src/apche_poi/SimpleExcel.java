package apche_poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleExcel {
	public static void main(String args[]) throws IOException,NullPointerException
	{
		String excelPath="D:/Alok/data.xlsx";
		FileInputStream inputStream=new FileInputStream(excelPath);// java class for taking externel file
		Workbook wb=new XSSFWorkbook(inputStream);// Workbook complete excel sheet
		Sheet firstSheet= wb.getSheetAt(0);// Per Sheet
		//Iterator<Row> iterator = firstSheet.iterator();
		int i=1;
		while(i<=4)
		{
			Row nextRow=firstSheet.getRow(i);// 
			//Cell cell=nextRow.getCell(i);
			//Iterator<Cell> cellIterator = nextRow.cellIterator();
			System.out.print(nextRow.getCell(0));
			System.out.print(nextRow.getCell(1));
			/*
			while(cellIterator.hasNext())
			{
				Cell cell=cellIterator.next();
				switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    System.out.print(cell.getStringCellValue()+"  ");
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    System.out.print(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    System.out.print(cell.getNumericCellValue());
                    break;
					}
				
			}*/
			System.out.println();
			i++;
		}
		
		
	}

}
