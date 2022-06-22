package com.jaggerabney.deviceaudit;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) {
        try {
            // load first sheet in frmInventory.xlsx
            Sheet sheet = loadWorkbook("frmInventory.xlsx").getSheetAt(0);
            String[] assetTags = getValuesOfColumn(sheet, getIndexOfColumn(sheet, "AssetTag"));
            String[] serialNums = getValuesOfColumn(sheet, getIndexOfColumn(sheet, "SerialNum"));
            String[] modelNames = getValuesOfColumn(sheet, getIndexOfColumn(sheet, "ItemDesc"));
            String[] statuses = getValuesOfColumn(sheet, getIndexOfColumn(sheet, "status"));
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private static XSSFWorkbook loadWorkbook(String filename) throws IOException {
        InputStream is = App.class.getClassLoader().getResourceAsStream(filename);
        return new XSSFWorkbook(is);
    }

    private static int getIndexOfColumn(Sheet sheet, String columnHeader) {
        int numCol = sheet.getRow(0).getPhysicalNumberOfCells();
        int colIndex = -1;
        String temp = "";
        DataFormatter df = new DataFormatter();

        for (int i = 0; i < numCol; i++) {
            temp = df.formatCellValue(sheet.getRow(0).getCell(i));

            if (temp.equals(columnHeader)) {
                colIndex = i;
            }
        }

        return colIndex;
    }

    private static String[] getValuesOfColumn(Sheet sheet, int colIndex) {
        ArrayList<String> tempValues = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        Cell currentCell = null;

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            currentCell = sheet.getRow(i).getCell(colIndex);
            tempValues.add(df.formatCellValue(currentCell));
        }

        return tempValues.toArray(new String[0]);
    }

}
