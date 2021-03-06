package com.company;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {

    public final static String cme_url = "http://www.cmegroup.com/CmeWS/exp/voiProductsViewExport.ctl?media=xls&tradeDate=20180410&assetClassId=3&reportType=P&excluded=CEE,CEU,KCB";
    public final static String primary_file = "/Users/bimby//IdeaProjects//Excel_cut/VoiTotalsByAssetClassExcelExport.xls";
    public final static String trimmed_file = "/Users/bimby//IdeaProjects//Excel_cut/VoiTotalsByAssetClassExcelExport_1.xls";
    public final static String converted_file = "/Users/bimby//IdeaProjects//Excel_cut/VoiTotalsByAssetClassExcelExport_2.txt";

    public List<List<HSSFCell>> cellGrid;

    public static void main(String[] args) throws IOException {
        Main obj = new Main();
        obj.onDownload();

        obj.convertExcelToTxt();
        obj.Futures_OI();
        obj.AudUsd_Options();
        obj.EurUsd_Options();
        obj.GbpUsd_Options();
        obj.UsdCad_Options();
        obj.UsdJpy_Options();
        obj.UsdChf_Options();

    }

    public void onDownload() {
        try {
            Download(cme_url, primary_file);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void Download(String url_Str, String primary_file) throws IOException {
        URL cme_url = new URL(url_Str);
        BufferedInputStream bInputStream = new BufferedInputStream(cme_url.openStream());
        FileOutputStream fOutputStream = new FileOutputStream(primary_file);
        byte[] buffer = new byte[1024];
        int count = 0;
        while ((count = bInputStream.read(buffer, 0, 1024)) != -1) {
            fOutputStream.write(buffer, 0, count);
        }
        fOutputStream.close();
        bInputStream.close();
    }

    public void convertExcelToTxt() throws IOException {
        try {
            cellGrid = new ArrayList<List<HSSFCell>>();
            FileInputStream myInput = new FileInputStream(primary_file);
            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
            HSSFSheet mySheet = myWorkBook.getSheetAt(0);
            Iterator<?> rowIter = mySheet.rowIterator();

            while (rowIter.hasNext()) {
                HSSFRow myRow = (HSSFRow) rowIter.next();
                Iterator<?> cellIter = myRow.cellIterator();
                List<HSSFCell> cellRowList = new ArrayList<HSSFCell>();
                while (cellIter.hasNext()) {
                    HSSFCell myCell = (HSSFCell) cellIter.next();
                    cellRowList.add(myCell);
                }
                cellGrid.add(cellRowList);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        File file = new File(converted_file);
        PrintStream stream = new PrintStream(file);
        for (int i = 0; i < cellGrid.size(); i++) {
            List<HSSFCell> cellRowList = cellGrid.get(i);
            for (int j = 0; j < cellRowList.size(); j++) {
                HSSFCell myCell = (HSSFCell) cellRowList.get(j);
                String stringCellValue = myCell.toString();
                stream.print(stringCellValue + "|");
            }
            stream.println("");
        }
    }

    public void Futures_OI() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Australian Dollar Future"))
                System.out.println("AudUsd Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Euro FX Future"))
                System.out.println("EurUsd Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("British Pound Future"))
                System.out.println("GbpUsd Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Canadian Dollar Future"))
                System.out.println("UsdCad Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Japanese Yen Future"))
                System.out.println("UsdJpy Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("New Zealand Dollar Future"))
                System.out.println("NzdUsd Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));

            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Swiss Franc Future"))
                System.out.println("UsdChf Futures " + sheet.getRow(i).getCell(2) + " VOI " + sheet.getRow(i).getCell(6));
        }
    }

    public void AudUsd_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);


        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Australian Dollar/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("AudUsd Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Australian Dollar/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("AudUsd Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }

    public void EurUsd_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Euro/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("EurUsd Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Euro/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("EurUsd Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }

    public void GbpUsd_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on British Pound/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("GbpUsd Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on British Pound/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("GbpUsd Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }

    public void UsdCad_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Canadian Dollar/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("UsdCad Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Canadian Dollar/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("UsdCad Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }

    public void UsdJpy_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Japanese Yen/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("UsdJpy Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Japanese Yen/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("UsdJpy Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }

    public void UsdChf_Options() throws IOException {

        FileInputStream myInput = new FileInputStream(primary_file);
        POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
        HSSFWorkbook workbook = new HSSFWorkbook(myFileSystem);

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Swiss Franc/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Calls")) {
                System.out.println("UsdChf Call Option " + sheet.getRow(i).getCell(2));
            }
            if (sheet.getRow(i) != null && sheet.getRow(i).getCell(0).getStringCellValue().equals("Premium Quoted European Style Options on Swiss Franc/US Dollar Future") && sheet.getRow(i).getCell(1).getStringCellValue().equals("Puts")) {
                System.out.println("UsdChf Put Option " + sheet.getRow(i).getCell(2));
            }
        }
    }
}
