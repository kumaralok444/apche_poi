package apche_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2 {
	public static void main(String args[]) throws IOException
	{
		String excelPath="D:/Alok/data.xlsx";
		FileInputStream is=new FileInputStream(excelPath);
		Workbook wb=new XSSFWorkbook(is);
		Sheet sh=wb.getSheetAt(0);
		int rn=sh.getLastRowNum();
		int i=1;
		while(i<=rn)
		{
			Row r=sh.getRow(i);
			int cn=r.getLastCellNum();
			System.out.println(cn);
			int rn1=r.getRowNum();
			System.out.println(rn1);
			int j=0;
			while(j<cn)
				{
					Cell ce=r.getCell(j);
					System.out.println(ce);
					j++;
				}
			i++;
		}
	}

}
