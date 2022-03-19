package com.sgtesting.excelfileassignments;


	import java.io.File;
	import java.io.FileOutputStream;

	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class FlowerNames 
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
				File f = new File("F:\\Flowers");
				boolean a = f.mkdir();
				System.out.println(a);
				File f1 = new File("F:\\Flowers\\Flowers.xlsx");
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
				row = sh.createRow(9);
				cell = row.createCell(0);
				cell.setCellValue("Flower 1");
				cell = row.createCell(1);
				cell.setCellValue("Flower 2");
				cell = row.createCell(2);
				cell.setCellValue("Flower 3");
				cell = row.createCell(3);
				cell.setCellValue("Flower 4");
				cell = row.createCell(4);
				cell.setCellValue("Flower 5");
				cell = row.createCell(5);
				cell.setCellValue("Flower 6");
				cell = row.createCell(6);
				cell.setCellValue("Flower 7");
				cell = row.createCell(7);
				cell.setCellValue("Flower 8");
				cell = row.createCell(8);
				cell.setCellValue("Flower 9");
				cell = row.createCell(9);
				cell.setCellValue("Flower 10");
				cell = row.createCell(10);
				cell.setCellValue("Flower 11");
				cell = row.createCell(11);
				cell.setCellValue("Flower 12");
				cell = row.createCell(12);
				cell.setCellValue("Flower 13");
				cell = row.createCell(13);
				cell.setCellValue("Flower 14");
				cell = row.createCell(14);
				cell.setCellValue("Flower 15");
				cell = row.createCell(15);
				cell.setCellValue("Flower 16");
				cell = row.createCell(16);
				cell.setCellValue("Flower 17");
				cell = row.createCell(17);
				cell.setCellValue("Flower 18");
				cell = row.createCell(18);
				cell.setCellValue("Flower 19");
				cell = row.createCell(19);
				cell.setCellValue("Flower 20");
				
				fout=new FileOutputStream("F:\\Flowers\\Flowers.xlsx");
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
