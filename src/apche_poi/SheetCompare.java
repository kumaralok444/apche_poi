package pck1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetCompare {
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		Map<Integer,String> colName=new HashMap<Integer,String>();
		colName.put(0, "Roll");
		colName.put(1, "Name");
		colName.put(2, "Class");
		String result="";
		String remark="Unamtched Column";
		String f="C:\\Users\\prasanth\\Desktop\\Test.xlsx";
		FileInputStream fi=new FileInputStream(f);
		Workbook wb =new XSSFWorkbook(fi);
		Sheet sh=wb.getSheet("Sheet1");
		Sheet sh1=wb.getSheet("Sheet2");
		Cell c=sh1.getRow(0).createCell(3);
		c.setCellValue("Result");
		c=sh1.getRow(0).createCell(4);
		c.setCellValue("Remark");
		//sheet2
		int sheet2RowNum=sh1.getLastRowNum();
		for(int i=1;i<=sheet2RowNum;i++) {
			String rollinSheet2=sh1.getRow(i).getCell(0).toString();
			result="PASS";
			remark="Unamtched Column: ";
		//sheet1
			int sheet1RowNum=sh.getLastRowNum();
			for(int j=1;j<=sheet1RowNum;j++) {
				String rollinSheet1=sh.getRow(j).getCell(0).toString();
				
				
				if(rollinSheet2.equalsIgnoreCase(rollinSheet1)) {
					int lastCellnum=sh.getRow(j).getLastCellNum();
					for(int k=1;k<lastCellnum;k++) {
						String stInSheet2=sh1.getRow(i).getCell(k).toString();
						String stInSheet1=sh.getRow(j).getCell(k).toString();
						if(!stInSheet2.equalsIgnoreCase(stInSheet1)) {
							//System.out.println("FAIL");
							remark=remark+colName.get(k);
							result="Fail";
						}
					}
					//System.out.println("PASS");
					break;
				}
				if(j==sheet1RowNum) {
					//System.out.println("data not found");
					result="Fail";
					remark="Roll Not Found";
				}
			}
			if(result.equalsIgnoreCase("PASS")) {
				c=sh1.getRow(i).createCell(3);
				c.setCellValue("PASS");
			}
			else {
				c=sh1.getRow(i).createCell(3);
				c.setCellValue("Fail");
				c=sh1.getRow(i).createCell(4);
				c.setCellValue(remark);
			}
		}
		FileOutputStream fos=new FileOutputStream(f);
		wb.write(fos);
		fos.close();
	}
}
