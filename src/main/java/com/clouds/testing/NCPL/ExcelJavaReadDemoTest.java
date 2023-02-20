package com.clouds.testing.NCPL;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.security.Key;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonObject;

public class ExcelJavaReadDemoTest {
// private static final String name = "/home/developers/Downloads/file.xlsx";
	static JsonObject object = new JsonObject();
//	static String key1;
//	static String key2;
//	static String key3;
//	static String key4;
	static Workbook workbook;
	static int columns1, columns2, columns3;
	static String sheetName;
	static Sheet sh;

	public static void main(String[] args) throws FileNotFoundException {
		Scanner sc = new Scanner(System.in);
// Map<String, Value> studentData = new TreeMap<String, Value>();
		try {
			ArrayList<String> list = new ArrayList<String>();
			FileInputStream file = new FileInputStream(new File("/home/developers/Downloads/file.xlsx"));
			workbook = new XSSFWorkbook(file);
			DataFormatter dataformatter = new DataFormatter();

			Iterator<Sheet> sheets = workbook.sheetIterator();
			int sheetno = 0;
			while (sheets.hasNext()) {
				sh = sheets.next();
				sheetName = sh.getSheetName();
				Iterator<Row> rowIterator = sh.iterator();
				columns1 = 0;
				while (rowIterator.hasNext()) {
					JsonObject object = new JsonObject();
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.iterator();
					{
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();

							String cellValue = dataformatter.formatCellValue(cell);
							if (row.getRowNum() == 0 && !(cellValue.isEmpty())) {
								list.add(cellValue);
							} else {
								if (row.getRowNum() != 0)
									columns1 = cell.getColumnIndex();
								switch (sheetno) {
//								for sheet no 1
								case 0: {
									object.addProperty(list.get(columns1), cellValue);
								}
									break;
//									for sheet no 2
								case 1: {
									if (row.getRowNum() != 0)
										columns2 = 4 + cell.getColumnIndex();
									object.addProperty(list.get(columns2), cellValue);
								}
									break;
//									for 3rd sheet
								case 2: {
									if (row.getRowNum() != 0)
										columns2 = 8 + cell.getColumnIndex();
									object.addProperty(list.get(columns2), cellValue);
								}
								
									break;

								default:
									break;
								}
							}
						}
						if (row.getRowNum() != 0)
							System.out.print(object);

						else {
							System.out.println(sh.getSheetName());
						}
					}
					System.out.println();

				}
				sheetno++;
//				System.out.println(sheetno);
			}

//			System.out.println("Choose sheet no");
//			int shno = sc.nextInt();
//
//			if (shno <= 3)
//				m1(workbook, sh, shno);

			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	

//		to get record sheetwise
//	public static void m1(Workbook workbook, Sheet sh, int shno) throws IOException {
//
//		ArrayList<String> list1 = new ArrayList<String>();
//
//		DataFormatter dataformatter = new DataFormatter();
//		Iterator<Sheet> sheets = workbook.sheetIterator();
//		int sheetCount = 0;
//		while (sheets.hasNext()) {
//			sh = sheets.next();
//			Iterator<Row> rowIterator = sh.iterator();
//			if (--shno == sheetCount) {
//				while (rowIterator.hasNext()) {
//					JsonObject object = new JsonObject();
//					Row row = rowIterator.next();
//					Iterator<Cell> cellIterator = row.iterator();
//					{
//						while (cellIterator.hasNext()) {
//							Cell cell = cellIterator.next();
//
//							String cellValue = dataformatter.formatCellValue(cell);
//							if (row.getRowNum() == 0 && !(cellValue.isEmpty())) {
//								list1.add(cellValue);
//							} else {
//								if (row.getRowNum() != 0) {
//									columns1 = cell.getColumnIndex();
//									object.addProperty(list1.get(columns1), cellValue);
//								}
//							}
//						}
//					}
//
//					System.out.println(object);
//				}
//			}
//			sheetCount++;
////				System.out.println(sheetno);
//			
//		}
//		System.out.println();
//		workbook.close();
//
//	}
}
