package com.bighalf.excel.randomizer.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Workbook;

public class Utilities {
	public static void closeStreams(Workbook wb, FileOutputStream fileOut,
			FileInputStream fin) {
		if (wb != null) {
			try {
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if (fileOut != null) {
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if (fin != null) {
			try {
				fin.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	
//	public static double changeValue(Double value, Double percentage) {
//		return value+=value*(percentage/100.00);
//    }
	
	public static BigDecimal changeValue(BigDecimal value, BigDecimal percentage) {
		return value.add(value.multiply(percentage.divide(new BigDecimal(100))));
    }
	
	
	public static void main(String[] args) {
		BigDecimal value = new BigDecimal(123);
		System.out.println(changeValue(value,new BigDecimal(10)));
		System.out.println(changeValue(new BigDecimal(135),new BigDecimal(-10)));
	}
}
