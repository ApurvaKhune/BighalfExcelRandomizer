package com.bighalf.excel.randomizer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.bighalf.excel.randomizer.util.Utilities;

public class Randomizer {

	private static Randomizer randomizer = null;
	static {
		randomizer = new Randomizer();
		System.out.println("1. argument must be file");
		System.out.println("2. argument must be the percentage ");
		System.out
				.println("EG : java -jar ExcelRandomizer.jar D:/input.xlsx 10");
		System.out
				.println("EG : java -jar ExcelRandomizer.jar D:/input.xlsx -10");
	}

	// below arguments represents the user input
	String pathname;
	double percentage;
	boolean incrementFlag;
	int userArgumentsLength = 2;// 1st argument must be the file path and
								// 2nd must be the percentage
								// 3rd must be increment flag
	// below arguments are used in the excel parsing
	File file;
	Workbook wb;
	FileOutputStream fileOut;
	FileInputStream fin;
	String extension = "";
	public static void main(String[] args) {

		String extension = randomizer.getUserInputs(args).getFile()
				.getExtension();
		randomizer.readFile(extension);
	}

	public Randomizer getUserInputs(String[] args) {
		if (args.length != userArgumentsLength) {
			System.out.println("only "+userArgumentsLength+" arguments are accepted");
			System.exit(0);
		}
		this.pathname = args[0];
		try {
			this.percentage = Double.parseDouble(args[1]);
			if (this.percentage == 0) {
				System.out.println("percentage cannot be zero");
				System.exit(0);
			}
		} catch (Exception e) {
			System.out.println("second argument must be an integer");
			System.exit(0);
		}
		return randomizer;
	}

	public Randomizer getFile() {
		File file = new File(this.pathname);
		if (!file.exists()) {
			System.out.println("file doesn't exists.");
			System.exit(0);
		}
		this.file = file;
		return randomizer;
	}

	public String getExtension() {
		
		String fileName = this.file.getName();
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
			this.extension = fileName.substring(i + 1);
		}
		return this.extension;
	}

	public String getResultExtension() {
		String fileName = this.file.getName();
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
			fileName = fileName.substring(0,i)+"_result."+this.extension;
		}
		return fileName;
	}

	public void readFile(String extension) {
		switch (extension) {
		case "xlsx":
			readExcelFile("xlsx");
			break;
			
		case "xls":
			readExcelFile("xls");
			break;

		default:
			System.out.println("only xlsx and xls format is supported");
			System.exit(0);
			break;
		}
	}

	public void readExcelFile(String fileExtension) {
		try {
			fin = new FileInputStream(file);
			if(fileExtension.equals("xlsx")){
				wb = new XSSFWorkbook(fin);
			}else{
				wb = new HSSFWorkbook(fin);
			}
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				Sheet s = wb.getSheetAt(i);
				for (Row row : s) {
					for (Cell cell : row) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							if (!DateUtil.isCellDateFormatted(cell)) {
								double cellValue = cell.getNumericCellValue();
								BigDecimal newCellValue = Utilities.changeValue(
										new BigDecimal(cellValue), new BigDecimal(this.percentage));
								cell.setCellValue(newCellValue.doubleValue());
							}
							break;
						}
					}
				}
			}
			fileOut = new FileOutputStream(getResultExtension());
			wb.write(fileOut);
			System.out.println("done");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			Utilities.closeStreams(wb, fileOut, fin);
		}
	}

}
