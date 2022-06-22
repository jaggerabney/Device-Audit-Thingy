package com.jaggerabney.deviceaudit;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

import com.jaggerabney.Device;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) {
        try {
            Workbook audit = loadWorkbook("audit.xlsx");
            Sheet sheet = audit.getSheet("Kellogg");
            System.out.println(Arrays.toString(getValuesOfColumn(sheet, getIndexOfColumn(sheet, "Tag #"))));

            // Device[] devices = createDevicesFromWorkbook(audit);

            // load first sheet in frmInventory.xlsx
            // Sheet frmInventory = loadWorkbook("frmInventory.xlsx").getSheetAt(0);
            // String[] frmAssetTags = getValuesOfColumn(frmInventory,
            // getIndexOfColumn(frmInventory, "AssetTag"));
            // String[] serialNums = getValuesOfColumn(frmInventory,
            // getIndexOfColumn(frmInventory, "SerialNum"));
            // String[] modelNames = getValuesOfColumn(frmInventory,
            // getIndexOfColumn(frmInventory, "ItemDesc"));
            // String[] statuses = getValuesOfColumn(frmInventory,
            // getIndexOfColumn(frmInventory, "Status"));

            // Workbook bowes62 = loadWorkbook("BOWES 6-2.xlsx");
            // fillValuesForColumn("AssetTag", assetTags, bowes62.getSheetAt(0));

            // OutputStream os = new FileOutputStream("BOWES 6-2.xlsx");
            // bowes62.write(os);
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private static XSSFWorkbook loadWorkbook(String filename) throws Exception {
        InputStream is = new FileInputStream(filename);
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
            // a blank row is considered null, hence this check
            if (sheet.getRow(i) != null) {
                currentCell = sheet.getRow(i).getCell(colIndex);

                if (currentCell == null || currentCell.getCellType() == CellType.BLANK) {
                    continue;
                }

                tempValues.add(df.formatCellValue(currentCell));
            }
        }

        return tempValues.toArray(new String[0]);
    }

    private static void fillValuesForColumn(String targetColumn, String[] columnData, Sheet sheet) {
        int targetColIndex = getIndexOfColumn(sheet, targetColumn);

        for (int i = 0; i < columnData.length; i++) {
            if (sheet.getRow(i) == null) {
                sheet.createRow(i);
            }

            sheet.getRow(i).createCell(targetColIndex).setCellValue(columnData[i]);
        }
    }

    private static Device[] createDevicesFromWorkbook(Workbook workbook) {
        ArrayList<Device> tempDevices = new ArrayList<>();
        int numSheets = workbook.getNumberOfSheets();
        Sheet currentSheet = null;
        String[] tempAssetTags = null;

        for (int i = 0; i < numSheets; i++) {
            currentSheet = workbook.getSheetAt(i);
            tempAssetTags = getValuesOfColumn(currentSheet, getIndexOfColumn(currentSheet, "Tag #"));

            for (int j = 0; j < tempAssetTags.length; j++) {

            }
        }

        return null;
    }
}
