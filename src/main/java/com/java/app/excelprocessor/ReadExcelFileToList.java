package com.java.app.excelprocessor;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.java.app.model.Country;

public class ReadExcelFileToList {

    public static void main(String[] args){
    List<Country> countryList = readExcelFileData("./excel/CountryList.xlsx");
    writeExcelFileData("./excel/CountryList.xlsx");
    System.out.println("Country List\n" + countryList);
}
    // reads excel file, processes data, creates new excel file
    private static void writeExcelFileData(String fileName){
        List<Country> countryList = readExcelFileData(fileName);
        String newFileName = fileName.replace(".xlsx", "Processed.xlsx");
        countryList.forEach(ReadExcelFileToList::process);

        try {
            FileOutputStream fos = new FileOutputStream(newFileName);
            XSSFWorkbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet();

            // process and populate the countries
            {
                int j = 0;
                Country c = null;
                for (Iterator<Country> i = countryList.iterator(); i.hasNext(); j++) {
                    c = i.next();
                    Row row = sheet.createRow(j);
                    Cell cell = row.createCell(0);
                    cell.setCellValue(c.getShortCode());
                    cell = row.createCell(1);
                    cell.setCellValue(c.getName());
                    cell = row.createCell(2);
                    cell.setCellValue(c.getTravelfriendly());
                };
            }    
            
            wb.write(fos);
            wb.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void process(Country c){
        int random = ThreadLocalRandom.current().nextInt(0, 3);
        switch(random){
            case 0:
                c.setTravelfriendly(Country.TRAVEL_FRIENDLY);
            break;
            case 1:
                c.setTravelfriendly(Country.RESTRICTED_TRAVEL);   
            break;
            case 2:
                c.setTravelfriendly(Country.QUARANTINE);
            break;
            default:
            throw new IllegalStateException("error in rndom generator");
        }
    }

    private static List<Country> readExcelFileData(String fileName){
        List<Country> countryList = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(fileName);
            XSSFWorkbook wb = new XSSFWorkbook(fis);

            int noOfSheets = wb.getNumberOfSheets();

            for (int i = 0; i < noOfSheets; i++) {
                Sheet sheet = wb.getSheetAt(i);
                Iterator<Row> rows = sheet.iterator();
                while(rows.hasNext()){
                    Row row = rows.next();
                    Iterator<Cell> cells = row.iterator();
                    Cell shortCodeC = row.getCell(0);
                    Cell nameC = row.getCell(1);
                    Cell serialNoC = row.getCell(2);
                    Country c = new Country(nameC.getStringCellValue().trim(),shortCodeC.getStringCellValue().trim());
                    countryList.add(c);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return countryList;
    }

}

