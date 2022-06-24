package com.jaggerabney.deviceaudit;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class App {
    public static final int NUM_DEVICE_PROPS = 7;
    public static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("#.00");

    public static void main(String[] args) {
        try {
            System.out.print("Loading audit.xlsx...");
            Workbook audit = loadWorkbook("audit.xlsx");
            System.out.print("done!\nLoading inventory.xlsx...");
            Workbook inventory = loadWorkbook("inventory.xlsx");
            System.out.print("done!\nLoading target.xlsx...");
            Workbook target = loadWorkbook("target.xlsx");
            System.out.print("done!\n");
            Device[] devices = createDevicesFromAuditWorkbook(audit);

            devices = updateDevicesWithInventoryInfo(inventory, devices);
            target = updateTargetWorkbookWithDeviceInfo(target, devices);

            System.out.print("\nWriting to target.xlsx...");
            OutputStream os = new FileOutputStream("target.xlsx");
            target.write(os);
            System.out.print("done!");
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

    private static Device[] createDevicesFromAuditWorkbook(Workbook workbook) {
        ArrayList<Device> tempDevices = new ArrayList<>();
        int numSheets = workbook.getNumberOfSheets();
        int roomColIndex = 0, assetColIndex = 1, cotColIndex = 2;
        double totalDevices = 0, numDevicesCreated = 0;
        Cell currentRoom, currentAsset, currentCot, lastRoom, lastCot;
        String room, asset, cot;
        Sheet currentSheet = null;
        Row currentRow = null;
        DataFormatter df = new DataFormatter();

        for (int sheet = 0; sheet < numSheets; sheet++) {
            currentSheet = workbook.getSheetAt(sheet);

            for (int row = 0; row < currentSheet.getPhysicalNumberOfRows(); row++) {
                if (currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(assetColIndex) != null) {
                    totalDevices++;
                }
            }
        }

        System.out.print(
                "Creating devices from audit workbook: " + (int) numDevicesCreated + " / " + (int) totalDevices + "("
                        + DECIMAL_FORMAT.format(((numDevicesCreated / totalDevices) * 100)) + "%)");

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
                    numDevicesCreated++;
                    System.out.print("\rCreating devices from audit workbook: " + (int) numDevicesCreated + " / "
                            + (int) totalDevices + " ("
                            + DECIMAL_FORMAT.format(((numDevicesCreated / totalDevices) * 100))
                            + "%)");
                }
            }
        }

        return tempDevices.toArray(new Device[0]);
    }

    private static Device[] updateDevicesWithInventoryInfo(Workbook workbook, Device[] devices) {
        Sheet sheet = workbook.getSheetAt(0);
        int assetColIndex = getIndexOfColumn(sheet, "AssetTag"),
                serialColIndex = getIndexOfColumn(sheet, "SerialNum"),
                modelColIndex = getIndexOfColumn(sheet, "ItemDesc"),
                statusColIndex = getIndexOfColumn(sheet, "Status");
        double numDevicesTotal = devices.length, numDevicesUpdated = 0;
        String currentSerial = "", currentModel = "", currentStatus = "";
        List<String> assetTags = Arrays.asList(getValuesOfColumn(sheet, assetColIndex));
        ArrayList<Device> result = new ArrayList<>();
        DataFormatter df = new DataFormatter();
        int currentRow = 0;

        System.out.print("\nUpdating devices with info from inventory.xlsx: " + (int) numDevicesUpdated + " / "
                + (int) numDevicesTotal + " (" + DECIMAL_FORMAT.format(((numDevicesUpdated / numDevicesTotal) * 100))
                + "%)");

        for (Device device : devices) {
            currentRow = assetTags.indexOf(device.asset);

            if (currentRow != -1) {
                currentSerial = df.formatCellValue(sheet.getRow(currentRow).getCell(serialColIndex));
                currentModel = df.formatCellValue(sheet.getRow(currentRow).getCell(modelColIndex));
                currentStatus = df.formatCellValue(sheet.getRow(currentRow).getCell(statusColIndex));

                result.add(
                        new Device(currentRow, device.asset, currentSerial, device.location, device.room, currentModel,
                                device.cot, currentStatus));
            }

            numDevicesUpdated++;
            System.out.print("\rUpdating devices with info from inventory.xlsx: " + (int) numDevicesUpdated + " / "
                    + (int) numDevicesTotal + " ("
                    + DECIMAL_FORMAT.format(((numDevicesUpdated / numDevicesTotal) * 100))
                    + "%)");

        }

        return result.toArray(new Device[0]);
    }

    private static Workbook updateTargetWorkbookWithDeviceInfo(Workbook workbook, Device[] devices) {
        Workbook result = workbook;
        Sheet sheet = result.getSheetAt(0);
        Row currentRow = null;
        Device currentDevice = null;
        int assetColIndex = getIndexOfColumn(sheet, "AssetTag"),
                serialColIndex = getIndexOfColumn(sheet, "Serial Number"),
                locationColIndex = getIndexOfColumn(sheet, "Location"),
                roomColIndex = getIndexOfColumn(sheet, "Room"),
                modelColIndex = getIndexOfColumn(sheet, "Model Name"),
                cotColIndex = getIndexOfColumn(sheet, "Checked Out to"),
                statusColIndex = getIndexOfColumn(sheet, "Status");
        double numRowsTotal = devices.length - 1, numRowsUpdated = 0; // -1 because both vars are 0-indexed

        System.out.print("\nUpdating target.xlsx with device info: " + (int) numRowsUpdated + " / " + (int) numRowsTotal
                + " (" + DECIMAL_FORMAT.format(((numRowsUpdated / numRowsTotal) * 100)) + "%)");

        for (int row = 1; row < devices.length; row++) {
            if (sheet.getRow(row) == null) {
                sheet.createRow(row);
            }
            currentRow = sheet.getRow(row);
            currentDevice = devices[row];

            for (int col = 0; col < NUM_DEVICE_PROPS; col++) {
                currentRow.createCell(col);
            }

            currentRow.getCell(assetColIndex).setCellValue(currentDevice.asset);
            currentRow.getCell(serialColIndex).setCellValue(currentDevice.serial);
            currentRow.getCell(locationColIndex).setCellValue(currentDevice.location);
            currentRow.getCell(roomColIndex).setCellValue(currentDevice.room);
            currentRow.getCell(modelColIndex).setCellValue(currentDevice.model);
            currentRow.getCell(cotColIndex).setCellValue(currentDevice.cot);
            currentRow.getCell(statusColIndex).setCellValue(currentDevice.status);

            numRowsUpdated++;
            System.out.print(
                    "\rUpdating target.xlsx with device info: " + (int) numRowsUpdated + " / " + (int) numRowsTotal
                            + " (" + DECIMAL_FORMAT.format(((numRowsUpdated / numRowsTotal) * 100)) + "%)");
        }

        return result;
    }

    private static boolean isEmpty(Cell cell) {
        return (cell == null || cell.getCellType() == CellType.BLANK);
    }
}
