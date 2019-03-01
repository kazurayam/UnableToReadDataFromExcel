package com.database

import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.annotation.Keyword

public class demo {

	XSSFWorkbook xssfWorkbook = null
	XSSFSheet xssfSheet = null

	@Keyword
	public String[][] readDataFromExcel(String filePath) throws IOException {
		String [][] testdata
		if (filePath.endsWith(".xlsx")) {
			File file = new File(filePath)
			println "file is ${file}"
			FileInputStream fis = new FileInputStream(file)
			xssfWorkbook = new XSSFWorkbook(fis)
			println "xssfWorkbook is ${xssfWorkbook}"
			
			xssfSheet = xssfWorkbook.getSheetAt(0);
			println "xssfSheet is ${xssfSheet}"
			
			//int numberOfRows = getRowCount()
			int numberOfRows = xssfSheet.getLastRowNum() + 1
			println "numberOfRows is ${numberOfRows}"
			int numberOfColumns = xssfSheet.getRow(1).getLastCellNum()
			println "numberOfColumns is ${numberOfColumns}"
			testdata = new String[numberOfRows - 1][numberOfColumns]
			for (int i = 1; i < numberOfRows; i++) {
				for (int j = 0; j < numberOfColumns; j++) {
					XSSFRow row = xssfSheet.getRow(i)
					XSSFCell cell = row.getCell(j)
					String value = xssfcellToString(cell)
					testdata[i - 1][j] = value
					if (value == null) {
						System.out.println("data empty")
					}
				}
			}
		}
		return testdata
	}


	public int getRowCount() {
		return xssfSheet.getLastRowNum() + 1
	}


	public String xssfcellToString(XSSFCell cell) {
		Object result=""
		if (cell != null) {
			int type = cell.getCellType()
			switch (type) {
				case 0:
					result = cell.getNumericCellValue()
					break
				case 1:
					result = cell.getStringCellValue()
					break
				case 2:
					result = cell.getCellFormula()
					break
				case 3:
					result = ""
					break
				default:
					throw new RuntimeException("no support for this cell")
			}
		}
		return result.toString()
	}
}
