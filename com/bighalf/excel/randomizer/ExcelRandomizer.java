//package com.bighalf.excel.randomizer;
//
//import java.io.BufferedInputStream;
//import java.io.BufferedOutputStream;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStream;
//import java.io.OutputStream;
//import java.io.PrintStream;
//import java.math.BigDecimal;
//import java.text.SimpleDateFormat;
//import java.util.Date;
//import java.util.HashMap;
//import java.util.LinkedHashMap;
//import org.apache.log4j.Logger;
//import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFDateUtil;
//import org.apache.poi.hssf.usermodel.HSSFRow;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFCellStyle;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
///*
// * This class specifies class file version 49.0 but uses Java 6 signatures.  Assumed Java 6.
// */
//public class ExcelRandomizer {
//    private static final Logger LOG = Logger.getLogger((Class)ExcelRandomizer.class);
//    private static Double percentageValue = 0.0;
//    private static int percentageType = 0;
//
//    public static void main(String[] args) {
//        block10 : {
//            if (args.length != 3) {
//                ExcelRandomizer.exitFromSystem("Insufficient Number of Arguments , # Arguments Requied : 1St Arg= File Full Path, 2nd Arg= percentage value , 3 Arg= 1 for increase 0 for descrese");
//            }
//            String FILE_PATH = args[0];
//            String percentage = args[1];
//            percentageValue = Double.valueOf(percentage);
//            percentageType = Integer.valueOf(args[2]);
//            if (percentageValue.doubleValue() == BigDecimal.ZERO.doubleValue()) {
//                ExcelRandomizer.exitFromSystem("Percentage must be  grater than 0  ");
//            }
//            if ("".equals(FILE_PATH) || FILE_PATH == null) {
//                ExcelRandomizer.exitFromSystem("File Path Should not be Empty");
//            } else {
//                FileInputStream fileStream = null;
//                try {
//                    File file = new File(FILE_PATH);
//                    fileStream = new FileInputStream(file);
//                    try {
//                        String fileType = ExcelRandomizer.getFileExtension(file);
//                        if (fileType.equalsIgnoreCase("xls")) {
//                            HSSFWorkbook workbook = new HSSFWorkbook((InputStream)new FileInputStream(file));
//                            System.out.println("Number of sheets :" + workbook.getNumberOfSheets());
//                            ExcelRandomizer.XLSFileProcess(file);
//                            break block10;
//                        }
//                        if (fileType.equalsIgnoreCase("xlsx")) {
//                            XSSFWorkbook workbook = new XSSFWorkbook((InputStream)new FileInputStream(file));
//                            System.out.println("Number of sheets :" + workbook.getNumberOfSheets());
//                            ExcelRandomizer.XLSXFileProcess(file);
//                            break block10;
//                        }
//                        throw new IllegalArgumentException("Received file does not have a standard excel extension.");
//                    }
//                    catch (IOException e) {
//                        e.printStackTrace();
//                    }
//                }
//                catch (FileNotFoundException e) {
//                    e.printStackTrace();
//                }
//            }
//        }
//    }
//
//    public static String getFileExtension(File file) {
//        String fileName = file.getName();
//        if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0) {
//            return fileName.substring(fileName.lastIndexOf(".") + 1);
//        }
//        return "";
//    }
//
//    public static void XLSFileProcess(File file) throws IOException {
//        System.out.println("\n Process is in Progress.... ##");
//        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
//        HSSFWorkbook currentWorkbook = new HSSFWorkbook((InputStream)bis);
//        HSSFWorkbook newWorkBook = new HSSFWorkbook();
//        HSSFSheet sheet = null;
//        HSSFRow row = null;
//        HSSFCell cell = null;
//        HSSFSheet mySheet = null;
//        HSSFRow myRow = null;
//        HSSFCell myCell = null;
//        int sheets = currentWorkbook.getNumberOfSheets();
//        int fCell = 0;
//        short lCell = 0;
//        int fRow = 0;
//        int lRow = 0;
//        for (int iSheet = 0; iSheet < sheets; ++iSheet) {
//            sheet = currentWorkbook.getSheetAt(iSheet);
//            if (sheet == null) continue;
//            mySheet = newWorkBook.createSheet(sheet.getSheetName());
//            fRow = sheet.getFirstRowNum();
//            lRow = sheet.getLastRowNum();
//            for (int iRow = fRow; iRow <= lRow; ++iRow) {
//                row = sheet.getRow(iRow);
//                myRow = mySheet.createRow(iRow);
//                if (row == null) continue;
//                fCell = row.getFirstCellNum();
//                lCell = row.getLastCellNum();
//                block10 : for (int iCell = fCell; iCell < lCell; ++iCell) {
//                    cell = row.getCell(iCell);
//                    myCell = myRow.createCell(iCell);
//                    if (cell == null) continue;
//                    myCell.setCellType(cell.getCellType());
//                    switch (cell.getCellType()) {
//                        case 3: {
//                            myCell.setCellValue("");
//                            continue block10;
//                        }
//                        case 4: {
//                            myCell.setCellValue(cell.getBooleanCellValue());
//                            continue block10;
//                        }
//                        case 5: {
//                            myCell.setCellErrorValue(cell.getErrorCellValue());
//                            continue block10;
//                        }
//                        case 2: {
//                            myCell.setCellFormula(cell.getCellFormula());
//                            continue block10;
//                        }
//                        case 0: {
//                            myCell.setCellValue(ExcelRandomizer.increase(cell.getNumericCellValue(), percentageValue, percentageType).doubleValue());
//                            continue block10;
//                        }
//                        case 1: {
//                            myCell.setCellValue(cell.getStringCellValue());
//                            continue block10;
//                        }
//                        default: {
//                            myCell.setCellFormula(cell.getCellFormula());
//                        }
//                    }
//                }
//            }
//        }
//        bis.close();
//        String outputFileName = "ExcelRandomizedCopy" + System.currentTimeMillis() + file.getName();
//        BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(outputFileName, true));
//        newWorkBook.write((OutputStream)bos);
//        bos.close();
//        currentWorkbook.close();
//        newWorkBook.close();
//        System.out.println("\n\n Process Ended.##\n\n OUTPUT FILE " + new File(outputFileName).getAbsolutePath());
//    }
//
//    public static void XLSXFileProcess(File file) throws IOException {
//        System.out.println("\n Process is in Progress.... ##");
//        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
//        XSSFWorkbook currentWorkbook = new XSSFWorkbook((InputStream)bis);
//        XSSFWorkbook newWorkBook = new XSSFWorkbook();
//        XSSFSheet sheet = null;
//        XSSFRow row = null;
//        XSSFCell cell = null;
//        XSSFSheet mySheet = null;
//        XSSFRow myRow = null;
//        XSSFCell myCell = null;
//        int sheets = currentWorkbook.getNumberOfSheets();
//        int fCell = 0;
//        short lCell = 0;
//        int fRow = 0;
//        int lRow = 0;
//        try {
//            for (int iSheet = 0; iSheet < sheets; ++iSheet) {
//                sheet = currentWorkbook.getSheetAt(iSheet);
//                if (sheet == null) continue;
//                mySheet = newWorkBook.createSheet(sheet.getSheetName());
//                fRow = sheet.getFirstRowNum();
//                lRow = sheet.getLastRowNum();
//                for (int iRow = fRow; iRow <= lRow; ++iRow) {
//                    row = sheet.getRow(iRow);
//                    myRow = mySheet.createRow(iRow);
//                    if (row == null) continue;
//                    fCell = row.getFirstCellNum();
//                    lCell = row.getLastCellNum();
//                    block12 : for (int iCell = fCell; iCell < lCell; ++iCell) {
//                        cell = row.getCell(iCell);
//                        myCell = myRow.createCell(iCell);
//                        if (cell == null) continue;
//                        myCell.setCellType(cell.getCellType());
//                        switch (cell.getCellType()) {
//                            case 3: {
//                                myCell.setCellValue("");
//                                continue block12;
//                            }
//                            case 4: {
//                                myCell.setCellValue(cell.getBooleanCellValue());
//                                continue block12;
//                            }
//                            case 5: {
//                                myCell.setCellErrorValue(cell.getErrorCellValue());
//                                continue block12;
//                            }
//                            case 2: {
//                                myCell.setCellFormula(cell.getCellFormula());
//                                continue block12;
//                            }
//                            case 0: {
//                                if (HSSFDateUtil.isCellDateFormatted((Cell)cell)) {
//                                    myCell.setCellValue(cell.toString());
//                                    continue block12;
//                                }
//                                myCell.setCellValue(ExcelRandomizer.increase(cell.getNumericCellValue(), percentageValue, percentageType).doubleValue());
//                                continue block12;
//                            }
//                            case 1: {
//                                myCell.setCellValue(cell.getStringCellValue());
//                                continue block12;
//                            }
//                            default: {
//                                myCell.setCellFormula(cell.getCellFormula());
//                            }
//                        }
//                    }
//                }
//            }
//        }
//        catch (Exception e) {
//            e.printStackTrace();
//        }
//        bis.close();
//        String outputFileName = "ExcelRandomizedCopy" + System.currentTimeMillis() + file.getName();
//        BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(outputFileName, true));
//        newWorkBook.write((OutputStream)bos);
//        bos.close();
//        currentWorkbook.close();
//        newWorkBook.close();
//        System.out.println("\n\n Process Ended.##\n\n OUTPUT FILE " + new File(outputFileName).getAbsolutePath());
//    }
//
//    public static Double increase(Double originalNumber, Double percentageValue, int percentageType) {
//        if (percentageType == 1) {
//            double percentage = percentageValue / 100.0;
//            double updatedValue = originalNumber * (1.0 + percentage);
//            return updatedValue;
//        }
//        double percentage = percentageValue / 100.0;
//        double updatedValue = originalNumber * (1.0 - percentage);
//        return updatedValue;
//    }
//
//    public LinkedHashMap<Integer, HashMap<String, String>> readData(XSSFWorkbook workbook) {
//        int i;
//        LinkedHashMap<Integer, HashMap<String, String>> rowMap = new LinkedHashMap<Integer, HashMap<String, String>>();
//        HashMap<Integer, String> column = new HashMap<Integer, String>();
//        if (workbook == null) {
//            return rowMap;
//        }
//        XSSFSheet sheet = workbook.getSheetAt(0);
//        if (sheet == null) {
//            return rowMap;
//        }
//        int flagRow = 0;
//        int flagCell = 0;
//        int totalRow = sheet.getLastRowNum();
//        for (i = 0; i < totalRow; ++i) {
//            if (sheet.getRow(i) == null) continue;
//            flagRow = i;
//            break;
//        }
//        for (i = 0; i < sheet.getRow(flagRow + 1).getLastCellNum(); ++i) {
//            if (sheet.getRow(flagRow + 1).getCell(i) == null) continue;
//            flagCell = i;
//            break;
//        }
//        int columnNumber = 1;
//        for (int i2 = flagCell; i2 < sheet.getRow(flagRow).getLastCellNum(); ++i2) {
//            if (sheet.getRow(flagRow).getCell(i2) == null) continue;
//            column.put(columnNumber, sheet.getRow(flagRow).getCell(i2).toString());
//            ++columnNumber;
//        }
//        HashMap<String, String> attributes = new HashMap<String, String>();
//        for (int j = 0; j < sheet.getRow(0).getLastCellNum(); ++j) {
//            if (sheet.getRow(0).getCell(j) == null) continue;
//            attributes.put(sheet.getRow(0).getCell(j).toString(), this.getCellType(sheet.getRow(1).getCell(j)));
//        }
//        int k = 1;
//        int i3 = 1;
//        for (i3 = 1; i3 <= totalRow; ++i3) {
//            if (sheet.getRow(i3) == null) continue;
//            HashMap<String, String> cellMap = new HashMap<String, String>();
//            for (int j2 = 0; j2 < sheet.getRow(i3).getLastCellNum(); ++j2) {
//                if (this.cellToString(sheet.getRow(0).getCell(j2)).equals("") && this.cellToString(sheet.getRow(i3).getCell(j2)).equals("")) continue;
//                cellMap.put(this.cellToString(sheet.getRow(0).getCell(j2)), this.cellToString(sheet.getRow(i3).getCell(j2)));
//            }
//            rowMap.put(k, cellMap);
//            ++k;
//        }
//        rowMap.put(0, attributes);
//        return rowMap;
//    }
//
//    public String getCellType(XSSFCell cell) {
//        String result = null;
//        switch (cell.getCellType()) {
//            case 0: {
//                result = HSSFDateUtil.isCellDateFormatted((Cell)cell) ? "Date" : "Numeric";
//                return result;
//            }
//            case 1: {
//                return "String";
//            }
//            case 4: {
//                return "Boolean";
//            }
//            case 2: {
//                return "Numeric";
//            }
//            case 3: {
//                return "String";
//            }
//            case 5: {
//                return "String";
//            }
//        }
//        System.out.print("unknown format");
//        return null;
//    }
//
//    public String cellToString(XSSFCell cell) {
//        String result = null;
//        if (cell == null) {
//            return "";
//        }
//        switch (cell.getCellType()) {
//            case 0: {
//                if (HSSFDateUtil.isCellDateFormatted((Cell)cell)) {
//                    if (cell.getCellStyle().getDataFormat() == 20) {
//                        SimpleDateFormat sdf = null;
//                        sdf = new SimpleDateFormat("HH:mm");
//                        Date date = cell.getDateCellValue();
//                        result = sdf.format(date);
//                    } else if (cell.getCellStyle().getDataFormat() == 17) {
//                        SimpleDateFormat sdf = null;
//                        sdf = new SimpleDateFormat("MM-yyyy");
//                        Date date = cell.getDateCellValue();
//                        result = sdf.format(date);
//                    } else if (cell.getCellStyle().getDataFormat() == 14) {
//                        SimpleDateFormat sdf = null;
//                        sdf = new SimpleDateFormat("MM-dd-yyyy");
//                        Date date = cell.getDateCellValue();
//                        result = sdf.format(date);
//                    }
//                } else {
//                    Object inputValue = null;
//                    long longVal = Math.round(cell.getNumericCellValue());
//                    double doubleVal = cell.getNumericCellValue();
//                    inputValue = Double.parseDouble(String.valueOf(longVal) + ".0") == doubleVal ? Long.valueOf(longVal) : Double.valueOf(doubleVal);
//                    result = inputValue.toString();
//                }
//                return result;
//            }
//            case 1: {
//                result = cell.getStringCellValue();
//                return result;
//            }
//            case 4: {
//                result = Boolean.toString(cell.getBooleanCellValue());
//                return result;
//            }
//            case 2: {
//                result = Double.toString(cell.getNumericCellValue());
//                return result;
//            }
//            case 3: {
//                result = "";
//                return result;
//            }
//            case 5: {
//                result = "ERROR";
//                return result;
//            }
//        }
//        result = "";
//        LOG.error((Object)"unknown format");
//        return null;
//    }
//
//    private static void exitFromSystem(String string) {
//        System.out.println("Error While Processing. \n Message  :  " + string);
//        System.exit(1);
//    }
//}