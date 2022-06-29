package com.jaggerabney.deviceaudit;

import java.io.*;
import java.text.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/*  TODO: 
 *    - talk with Kyle to get information from various PO/budgeting sheets
 */

public class App {
    public static final Scanner SCANNER = new Scanner(System.in);
    public static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("0.00");
    public static final DataFormatter DATA_FORMATTER = new DataFormatter();
    private static Properties PROPS;

    public static void main(String[] args) {
        // try block is used to catch IOException errors that come with reading
        // from/writing to files
        try {
            System.out.print("Loading config.properties...");
            PROPS = loadProps("config.properties");
            System.out.print("done!\n");

            File targetFile = new File(PROPS.getProperty("targetWorkbookName"));
            if (targetFile.exists()) {
                targetWorkbookExistsHandler();
            }

            // print statements signifying file loading are put outside of the loadWorkbook
            // method to better reflect the state of the program.
            // if they're inside the loadWorkbook method, then the program reports that it's
            // done loading files before it actually is, which makes it look like it's
            // crashing/hitching.

            // the below code "loads" - it's more accurate to say "gets references to" - all
            // necessary workbooks, then uses the audit.xlsx workbook to initialize an array
            // of Device objects. the Device object is just a simple object that keeps track
            // of all pertinent information that needs to be filled out in target.xlsx:
            // asset tag, serial number, model, location, what room it's in, etc.
            System.out.print("Loading audit workbook...");
            Workbook audit = loadWorkbook(PROPS.getProperty("auditWorkbookName"));
            System.out.print("done!\nLoading inventory workbook...");
            Workbook inventory = loadWorkbook(PROPS.getProperty("inventoryWorkbookName"));
            System.out.print("done!\n");
            // Workbook target = loadWorkbook(PROPS.getProperty("targetWorkbookName"));
            // System.out.print("done!\n");
            String location = askQuestion("What school is this sheet for?");
            Device[] devices = createDevicesFromAuditWorkbook(audit);

            // uses the x = change(x) method to get around the fact that java is
            // pass-by-value. the Device objects in the devices array initially only have
            // their asset, room, and cot (checked out to) fields written to at first. the
            // updateDevicesWithInventoryInfo pulls the rest of the needed info - S/N,
            // model, status, etc. - and rewrites all of the objects in the devices array
            // with these updated copies. see the Device file for more information
            devices = updateDevicesWithInventoryInfo(inventory, devices, location);
            // then, the target workbook (which is merely a Java representation of the
            // workbook, not the actual workbook itself) is updated so that the *actual*
            // target.xlsx workbook can be written to
            Workbook target = createTargetWorkbookWithDeviceInfo(devices);

            // writes to target.xlsx
            System.out.print("\nWriting to target workbook...");
            OutputStream os = new FileOutputStream(PROPS.getProperty("targetWorkbookName"));
            target.write(os);
            System.out.print("done!\nClosing target workbook...");
            target.close();
            System.out.print("closed!\n");

            SCANNER.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // loads props
    private static Properties loadProps(String filename) throws IOException {
        FileInputStream fis = new FileInputStream(filename);
        Properties props = new Properties();
        props.load(fis);
        fis.close();
        return props;

    }

    // small helper function for loading workbooks.
    // XSSFWorkbook is a concrete class, whereas Workbook is an interface
    private static XSSFWorkbook loadWorkbook(String filename) throws Exception {
        InputStream is = new FileInputStream(filename);
        return new XSSFWorkbook(is);
    }

    private static void targetWorkbookExistsHandler() {
        String answer = askQuestion(PROPS.getProperty("targetWorkbookExistsMessage"));
        if (answer.equalsIgnoreCase("Y")) {
            return;
        } else {
            System.exit(0);
        }
    }

    // helper function used to find the index of a column in a given sheet. since
    // the different worksheets have different names for the same thing - e.g. asset
    // tag may be written as "Asset Tag", "Asset", "Tag", "Tag #" - this function is
    // used to make the program a little more robust
    private static int getIndexOfColumn(Sheet sheet, String columnHeader) {
        int numCol = sheet.getRow(0).getPhysicalNumberOfCells();
        int colIndex = -1;
        String temp = "";

        // the actual function itself works by looping through the first row in the
        // given sheet and seeing if the text in each cell matches the text passed as an
        // argument to the function. simple, really
        for (int col = 0; col < numCol; col++) {
            temp = DATA_FORMATTER.formatCellValue(sheet.getRow(0).getCell(col));

            if (temp.equals(columnHeader)) {
                colIndex = col;
            }
        }

        return colIndex;
    }

    // helper function that returns the values of a given column in a sheet. useful
    // when, for example, you need all of the asset tags in inventory.xlsx.
    private static String[] getValuesOfColumn(Sheet sheet, int colIndex) {
        ArrayList<String> tempValues = new ArrayList<>();
        Cell currentCell = null;

        for (int row = 0; row < sheet.getPhysicalNumberOfRows(); row++) {
            // a blank row is considered null, hence this check
            if (sheet.getRow(row) != null) {
                currentCell = sheet.getRow(row).getCell(colIndex);

                // we don't want blank values in our returned array! so we skip em if they are
                if (isEmpty(currentCell)) {
                    continue;
                }

                tempValues.add(DATA_FORMATTER.formatCellValue(currentCell));
            }
        }

        return tempValues.toArray(new String[0]);
    }

    // function used to initialize the devices array in the main method
    private static Device[] createDevicesFromAuditWorkbook(Workbook workbook) {
        ArrayList<Device> tempDevices = new ArrayList<>();
        int numSheets = workbook.getNumberOfSheets();
        // hard-coded numbers are used here instead of the getIndexOfColumn method
        // because different techs named their columns different things. for example, i
        // tended to name my asset column "Asset", whereas other techs named theirs
        // "Tag" or "Tag #".
        int roomColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookRoomColIndex")),
                assetColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookAssetColIndex")),
                cotColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookUserColIndex"));
        double totalDevices = 0, numDevicesCreated = 0;
        Cell currentRoom, currentAsset, currentCot, lastRoom, lastCot;
        String room, asset, cot;
        Sheet currentSheet = null;
        Row currentRow = null;

        // the below for loop is used for the text that shows the progress in the
        // console. since each tech worked on their own sheet, the totality of devices
        // for a given school is spread across multiple sheets, hence why they are
        // looped through here. likewise, a for loop is used instead of Sheet's
        // getPhysicalNumberOfRows method to excise out blank rows
        for (int sheet = 0; sheet < numSheets; sheet++) {
            currentSheet = workbook.getSheetAt(sheet);

            for (int row = 1; row < currentSheet.getPhysicalNumberOfRows(); row++) {
                if (currentSheet.getRow(row) != null && currentSheet.getRow(row).getCell(assetColIndex) != null) {
                    totalDevices++;
                }
            }
        }

        System.out.print(
                "Creating devices from audit workbook: " + (int) numDevicesCreated + " / " + (int) totalDevices + "("
                        + DECIMAL_FORMAT.format(((numDevicesCreated / totalDevices) * 100)) + "%)");

        // loops through every sheet...
        for (int sheet = 0; sheet < numSheets; sheet++) {
            currentSheet = workbook.getSheetAt(sheet);
            lastRoom = null;
            lastCot = null;

            // ...then every row in the audit workbook to create an array of Device objects
            // based specifically on the asset column.
            // row is initialized to 1 to skip the header row
            for (int row = 1; row < currentSheet.getPhysicalNumberOfRows(); row++) {
                if (currentSheet.getRow(row) != null && !isEmpty(currentSheet.getRow(row).getCell(assetColIndex))) {
                    currentRow = currentSheet.getRow(row);
                    currentRoom = currentRow.getCell(roomColIndex);
                    currentAsset = currentRow.getCell(assetColIndex);
                    currentCot = currentRow.getCell(cotColIndex);

                    // if there is no specified room for a given row, the last specified room is
                    // used. this is what is referred to in the github repo as Room
                    // Interpolation(tm).
                    if (isEmpty(currentRoom)) {
                        currentRoom = lastRoom;
                        // also checks if the last specified room is different than the current room. if
                        // so, then the last specified user is no longer valid, and is gotten rid of.
                    } else if (lastRoom != null && lastCot != null && !lastRoom.equals(currentRoom)) {
                        lastCot.setBlank();
                    }
                    room = DATA_FORMATTER.formatCellValue(currentRoom);
                    asset = DATA_FORMATTER.formatCellValue(currentAsset);
                    // if there is no specified user for a given row, then the last specified user
                    // is used, like with the room. this is User Interpolation(tm).
                    if (isEmpty(currentCot)) {
                        currentCot = lastCot;
                    }
                    cot = DATA_FORMATTER.formatCellValue(currentCot);

                    // lastRoom/lastCot are set to the current room/cot for the next loop. this is
                    // the crux of User/Room Interpolation(tm).
                    lastRoom = currentRoom;
                    lastCot = currentCot;
                    // device is added to the temporary ArrayList, and the progress text is updated
                    // accordingly.

                    tempDevices.add(new Device(room, asset, cot, asset.trim().length() <= Integer
                            .valueOf(PROPS.getProperty("auditWorkbookAssetThreshold").trim())));
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

    // function used in main method to initialize the Device objects' fields that
    // weren't initialized in createDevicesFromAuditWorkbook
    private static Device[] updateDevicesWithInventoryInfo(Workbook workbook, Device[] devices, String location)
            throws Exception {
        Sheet sheet = workbook.getSheetAt(0);
        int assetColIndex = getIndexOfColumn(sheet, PROPS.getProperty("inventoryWorkbookAssetColName")),
                serialColIndex = getIndexOfColumn(sheet, PROPS.getProperty("inventoryWorkbookSerialColName")),
                modelColIndex = getIndexOfColumn(sheet, PROPS.getProperty("inventoryWorkbookModelColName")),
                statusColIndex = getIndexOfColumn(sheet, PROPS.getProperty("inventoryWorkbookStatusColName"));
        double numDevicesTotal = devices.length, numDevicesUpdated = 0;
        String currentAsset = "", currentSerial = "", currentModel = "", currentStatus = "";
        List<String> assetTags = Arrays.asList(getValuesOfColumn(sheet, assetColIndex));
        List<String> serialNums = Arrays.asList(getValuesOfColumn(sheet, serialColIndex));
        ArrayList<Device> result = new ArrayList<>();
        int currentRow = 0;
        boolean deviceHasAssetPopulated = false;

        // used for the progress text in the console
        System.out.print("\nUpdating devices with info from inventory workbook: " + (int) numDevicesUpdated + " / "
                + (int) numDevicesTotal + " (" + DECIMAL_FORMAT.format(((numDevicesUpdated / numDevicesTotal) * 100))
                + "%)");

        // the method loops through every device in devices and tries to find its
        // corresponding row in the inventory.xlsx workbook. since the assetTags
        // ArrayList<> consists of nothing but the values from inventory.xlsx, and since
        // inventory.xlsx has no blank rows in it, it's assumed that the index of an
        // asset tag in assetTags == the row that it's in inside inventory.xlsx.
        for (Device device : devices) {
            if (device.asset != null && device.serial == null) {
                deviceHasAssetPopulated = true;
                currentRow = assetTags.indexOf(device.asset);
            } else if (device.asset == null && device.serial != null) {
                deviceHasAssetPopulated = false;
                currentRow = serialNums.indexOf(device.serial);
            } else {
                throw new Exception("Invalid object state!");
            }

            // if currentRow == -1, then the device is *not* in inventory.xlsx. thus, this
            // information is not filled.
            if (currentRow != -1 && deviceHasAssetPopulated) {
                currentSerial = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(serialColIndex));
                currentModel = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(modelColIndex));
                currentStatus = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(statusColIndex));
            } else if (currentRow != -1 && !deviceHasAssetPopulated) {
                currentAsset = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(assetColIndex));
                currentModel = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(modelColIndex));
                currentStatus = DATA_FORMATTER.formatCellValue(sheet.getRow(currentRow).getCell(statusColIndex));
            } else if (currentRow == -1 && deviceHasAssetPopulated) {
                currentSerial = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
                currentModel = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
                currentStatus = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
            } else if (currentRow == -1 && !deviceHasAssetPopulated) {
                currentAsset = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
                currentModel = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
                currentStatus = PROPS.getProperty("cantFindAssetInInventoryWorkbookMessage");
            } else {
                throw new Exception("Invalid row state!");
            }

            if (deviceHasAssetPopulated) {
                result.add(new Device(device.asset, currentSerial, location, device.room, currentModel,
                        device.cot, currentStatus));
            } else {
                result.add(new Device(currentAsset, device.serial, location, device.room, currentModel,
                        device.cot, currentStatus));
            }

            // progress text in console is updated accordingly
            numDevicesUpdated++;
            System.out.print("\rUpdating devices with info from inventory workbook: " + (int) numDevicesUpdated + " / "
                    + (int) numDevicesTotal + " ("
                    + DECIMAL_FORMAT.format(((numDevicesUpdated / numDevicesTotal) * 100))
                    + "%)");

        }

        return result.toArray(new Device[0]);
    }

    // function used in main method to update target workbook object with info from
    // the devices array. note that this *doesn't* actually write to the target.xlsx
    // file; this just updates the corresponding object.
    private static Workbook createTargetWorkbookWithDeviceInfo(Device[] devices) {
        Workbook result = new XSSFWorkbook();
        Sheet sheet = result.createSheet("Data");
        Row currentRow = null;
        Device currentDevice = null;
        int assetColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookAssetColIndex")),
                serialColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookSerialColIndex")),
                locationColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookLocationColIndex")),
                roomColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookRoomColIndex")),
                modelColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookModelColIndex")),
                cotColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookUserColIndex")),
                statusColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookStatusColIndex"));
        String assetColName = PROPS.getProperty("targetWorkbookAssetColName"),
                serialColName = PROPS.getProperty("targetWorkbookSerialColName"),
                locationColName = PROPS.getProperty("targetWorkbookLocationColName"),
                roomColName = PROPS.getProperty("targetWorkbookRoomColName"),
                modelColName = PROPS.getProperty("targetWorkbookModelColName"),
                cotColName = PROPS.getProperty("targetWorkbookUserColName"),
                statusColName = PROPS.getProperty("targetWorkbookStatusColName");
        double numRowsTotal = devices.length, numRowsUpdated = 0;

        // writes progress text for first time in console
        System.out.print(
                "\nUpdating target workbook with device info: " + (int) numRowsUpdated + " / " + (int) numRowsTotal
                        + " (" + DECIMAL_FORMAT.format(((numRowsUpdated / numRowsTotal) * 100)) + "%)");

        // create headers
        Row headerRow = sheet.createRow(0);
        for (int col = 0; col < Device.NUM_PROPS; col++) {
            headerRow.createCell(col);
        }

        headerRow.getCell(assetColIndex).setCellValue(assetColName);
        headerRow.getCell(serialColIndex).setCellValue(serialColName);
        headerRow.getCell(locationColIndex).setCellValue(locationColName);
        headerRow.getCell(roomColIndex).setCellValue(roomColName);
        headerRow.getCell(modelColIndex).setCellValue(modelColName);
        headerRow.getCell(cotColIndex).setCellValue(cotColName);
        headerRow.getCell(statusColIndex).setCellValue(statusColName);

        // for loop starts at 1 because row 0 contains headers
        for (int row = 1; row <= devices.length; row++) {
            sheet.createRow(row);
            currentRow = sheet.getRow(row);
            currentDevice = devices[row - 1]; // because arrays are 0-indexed!

            // not only does a row need to be created in order to set the value of a cell,
            // but the actual cell itself needs to be created as well. otherwise, a
            // NullPointerException is thrown
            for (int col = 0; col < Device.NUM_PROPS; col++) {
                currentRow.createCell(col);
            }

            // ooga booga caveman brain solution, i know. i'd like to think that i'll
            // refactor this at some point
            currentRow.getCell(assetColIndex).setCellValue(currentDevice.asset);
            currentRow.getCell(serialColIndex).setCellValue(currentDevice.serial);
            currentRow.getCell(locationColIndex).setCellValue(currentDevice.location);
            currentRow.getCell(roomColIndex).setCellValue(currentDevice.room);
            currentRow.getCell(modelColIndex).setCellValue(currentDevice.model);
            currentRow.getCell(cotColIndex).setCellValue(currentDevice.cot);
            currentRow.getCell(statusColIndex).setCellValue(currentDevice.status);

            // updates progress text in console
            numRowsUpdated++;
            System.out.print(
                    "\rUpdating target workbook with device info: " + (int) numRowsUpdated + " / " + (int) numRowsTotal
                            + " (" + DECIMAL_FORMAT.format(((numRowsUpdated / numRowsTotal) * 100)) + "%)");
        }

        return result;
    }

    private static String askQuestion(String question) {
        System.out.print(question + " ");
        String result = SCANNER.nextLine();
        return result;
    }

    // helper function used to determine if a cell is empty or not
    private static boolean isEmpty(Cell cell) {
        return (cell == null || cell.getCellType() == CellType.BLANK);
    }
}
