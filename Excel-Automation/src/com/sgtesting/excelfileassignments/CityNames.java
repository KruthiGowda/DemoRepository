package com.sgtesting.excelfileassignments;

	import java.io.File;
	import java.io.FileOutputStream;

	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class CityNames  
	{
		public static void main(String[] args) 
		{
			createfolder();
			writecontent();
			
		}
		private static void createfolder()
		{
			try
			{
				File f = new File("F:\\City");
				boolean a = f.mkdir();
				System.out.println(a);
				File f1 = new File("F:\\City\\City.xlsx");
				boolean a1 = f1.createNewFile();
				System.out.println(a1);
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}
		private static void writecontent()
		{
			Workbook wb = null;
			Sheet sh = null;
			Row row = null;
			Cell cell = null;
			FileOutputStream fout = null;
			try
			{
				wb = new XSSFWorkbook();
				sh = wb.createSheet("Sheet 1");
				row = sh.createRow(4);
				cell = row.createCell(0);
				cell.setCellValue("City 1");
				cell = row.createCell(1);
				cell.setCellValue("City 2");
				cell = row.createCell(2);
				cell.setCellValue("City 3");
				cell = row.createCell(3);
				cell.setCellValue("City 4");
				cell = row.createCell(4);
				cell.setCellValue("City 5");
				cell = row.createCell(5);
				cell.setCellValue("City 6");
				cell = row.createCell(6);
				cell.setCellValue("City 7");
				cell = row.createCell(7);
				cell.setCellValue("City 8");
				cell = row.createCell(8);
				cell.setCellValue("City 9");
				cell = row.createCell(9);
				cell.setCellValue("City 10");
				cell = row.createCell(10);
				cell.setCellValue("City 11");
				cell = row.createCell(11);
				cell.setCellValue("City 12");
				cell = row.createCell(12);
				cell.setCellValue("City 13");
				cell = row.createCell(13);
				cell.setCellValue("City 14");
				cell = row.createCell(14);
				cell.setCellValue("City 15");
				cell = row.createCell(15);
				cell.setCellValue("City 16");
				cell = row.createCell(16);
				cell.setCellValue("City 17");
				cell = row.createCell(17);
				cell.setCellValue("City 18");
				cell = row.createCell(18);
				cell.setCellValue("City 19");
				cell = row.createCell(19);
				cell.setCellValue("City 20");
				
				fout=new FileOutputStream("F:\\City\\City.xlsx");
				wb.write(fout);
				
			}catch(Exception e)
			{
				e.printStackTrace();
			}
			finally
			{
				try
				{
					fout.close();
					wb.close();
				}catch(Exception e)
				{
					e.printStackTrace();
				}
			}
		}
}


