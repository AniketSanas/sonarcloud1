package com.clouds.testing.NCPL;

import java.io.*;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import com.google.gson.JsonObject;

public class ReadExcelJson2 {
	static String Emp_id;
	static String Emp_name;
	static String value1;
	static String value2;

	public static void main(String[] args) throws IOException {
//		File file = new File("");
		JsonObject mainObject=new JsonObject();

		// obtaining input bytes from a file
		FileInputStream fis = new FileInputStream(new File("/home/developers/Downloads/NCPL1Demo.xls"));
		// creating workbook instance that refers to .xls file
		System.out.println(fis);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		System.out.println(wb);
		// creating a Sheet object to retrieve the object
		
//		
		DataFormatter formatter = new DataFormatter();

		
		
		HSSFSheet sheet = wb.getSheetAt(0);
		System.out.println(sheet);
		// evaluating cell type
		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
		System.out.println(formulaEvaluator);
		int count = 0;
		for (Row row : sheet) // iteration over row using for each loop
		{
			int count1 = 0;
			//			System.out.println(row);
			for (Cell cell : row) // iteration over cell using for each loop
			{
				
				        CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
//				        System.out.print(cellRef.formatAsString());
				        String text = formatter.formatCellValue(cell);
//				        System.out.print(text+" ");
//				        System.out.println(cell.getCellType());
//				        if(cell.getCellType()==CellType.STRING)
//		                System.out.println(cell.getRichStringCellValue().getString());
//		                System.out.println(cell.getRichStringCellValue().getString());


				if (count == 0) {
					switch (count1) {
					case 0: {
						Emp_id = text;
//						System.out.println(Emp_id);
					}
						break;
					case 1: {
						Emp_name = text;
//						System.out.println(Emp_name);

					}
						break;
					default: {
						System.out.println("keys not stored in Emp_id,Emp_name");
					}
					}
				} else {
					switch (count1) {
					case 0: {
						
						value1=text;
					}
						break;
					case 1: {
						value2=text;
					}
						break;
					default: {
						System.out.println("keys not stored in Emp_id,Emp_name");
					}
					}
				}
				count1++;
			}
//			}
			if(count!=0)
			{
			mainObject.addProperty(Emp_id,value1);
			mainObject.addProperty(Emp_name,value2);
			System.out.println(mainObject);
			}

			count++;
//			System.out.println("count="+count);

		}
//		System.out.println("count="+count);
		

	}

}











































