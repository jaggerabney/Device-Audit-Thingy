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
            Device[] devices = createDevicesFromAuditWorkbook(audit);
            System.out.println(Arrays.toString(devices));

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

                if (isEmpty(currentCell)) {
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

    private static Device[] createDevicesFromAuditWorkbook(Workbook workbook) {
        ArrayList<Device> tempDevices = new ArrayList<>();
        int numSheets = workbook.getNumberOfSheets();
        int roomColIndex = 0, assetColIndex = 1, cotColIndex = 2;
        Cell currentRoom, currentAsset, currentCot, lastRoom, lastCot;
        String room, asset, cot;
        Sheet currentSheet = null;
        Row currentRow = null;
        DataFormatter df = new DataFormatter();

        for (int sheet = 0; sheet < numSheets; sheet++) {
            currentSheet = workbook.getSheetAt(sheet);
            lastRoom = null;
            lastCot = null;

            for (int row = 0; row < currentSheet.getPhysicalNumberOfRows(); row++) {
                if (currentSheet.getRow(row) != null) {
                    currentRow = currentSheet.getRow(row);
                    currentRoom = currentRow.getCell(roomColIndex);
                    currentAsset = currentRow.getCell(assetColIndex);
                    currentCot = currentRow.getCell(cotColIndex);

                    if (isEmpty(currentRoom)) {
                        currentRoom = lastRoom;
                    }
                    room = df.formatCellValue(currentRoom);
                    asset = df.formatCellValue(currentAsset);
                    if (isEmpty(currentCot)) {
                        currentCot = lastCot;
                    }
                    cot = df.formatCellValue(currentCot);

                    lastRoom = currentRoom;
                    lastCot = currentCot;
                    tempDevices.add(new Device(row, room, asset, cot));
                }
            }
        }

        return tempDevices.toArray(new Device[0]);
    }

    private static boolean isEmpty(Cell cell) {
        return (cell == null || cell.getCellType() == CellType.BLANK);
    }
}
