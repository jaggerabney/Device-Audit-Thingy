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
            String[] assetTags = getAssetTagsFrom(sheet);
            System.out.println(Arrays.toString(assetTags));
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private static XSSFWorkbook loadWorkbook(String filename) throws IOException {
        InputStream is = App.class.getClassLoader().getResourceAsStream(filename);
        return new XSSFWorkbook(is);
    }

    private static String[] getAssetTagsFrom(Sheet sheet) throws Exception {
        int numCol = sheet.getRow(0).getPhysicalNumberOfCells();
        int assetTagColIndex = -1;
        ArrayList<String> tempAssetTags = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        String temp = "";

        // find column that stores asset tags
        for (int i = 0; i < numCol; i++) {
            temp = df.formatCellValue(sheet.getRow(0).getCell(i));

            if (temp.replaceAll("\\s+", "").equals("AssetTag")) {
                assetTagColIndex = i;
            }
        }

        // throws exception if column can't be found
        if (assetTagColIndex == -1) {
            throw new Exception("Asset tag column not found!");
        }

        // stores value in asset tag column
        Cell currentCell = null;

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            currentCell = sheet.getRow(i).getCell(assetTagColIndex);
            tempAssetTags.add(df.formatCellValue(currentCell));
        }

        return tempAssetTags.toArray(new String[0]);
    }
}
