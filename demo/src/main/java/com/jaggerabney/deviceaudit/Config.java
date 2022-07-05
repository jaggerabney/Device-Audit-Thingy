package com.jaggerabney.deviceaudit;

import java.util.*;
import java.io.*;

public class Config {
        private static final Properties PROPS = loadProps("config.properties");

        // audit workbook stuff
        public static final int auditWorkbookRoomColIndex = Integer
                        .valueOf(PROPS.getProperty("auditWorkbookRoomColIndex")),
                        auditWorkbookAssetColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookAssetColIndex")),
                        auditWorkbookUserColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookUserColIndex")),
                        auditWorkbookAssetThreshold = Integer.valueOf(PROPS.getProperty("auditWorkbookAssetThreshold"));

        // workbook/sheet names
        public static final String auditWorkbookName = PROPS.getProperty("auditWorkbookName"),
                        inventoryWorkbookName = PROPS.getProperty("inventoryWorkbookName"),
                        targetWorkbookName = PROPS.getProperty("targetWorkbookName"),
                        targetWorkbookSheetName = PROPS.getProperty("targetWorkbookSheetName");

        // inventory workbook column names/indices
        public static final Col inventoryWorkbookAssetCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookAssetColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookAssetColIndex")));
        public static final Col inventoryWorkbookSerialCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookSerialColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookSerialColIndex")));
        public static final Col inventoryWorkbookModelCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookModelColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookModelColIndex")));
        public static final Col inventoryWorkbookStatusCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookStatusColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookStatusColIndex")));
        public static final Col inventoryWorkbookBudgetNumCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookBudgetNumColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookBudgetNumColIndex")));
        public static final Col inventoryWorkbookPurchDateCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookPurchDateName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookPurchDateColIndex")));
        public static final Col inventoryWorkbookPurchPriceCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookPurchPriceColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookPurchPriceColIndex")));
        public static final Col inventoryWorkbookPurchOrderNumCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookPurchOrderNumColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookPurchOrderNumColIndex")));
        public static final Col inventoryWorkbookProductNumCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookProductNumColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookProductNumColIndex")));
        public static final Col inventoryWorkbookModelNumCol = Col.newCol(
                        PROPS.getProperty("inventoryWorkbookModelNumColName"),
                        Integer.valueOf(PROPS.getProperty("inventoryWorkbookModelNumColIndex")));

        // target workbook column names (and indices i guess?)
        public static final List<String> targetWorkbookColumns = Arrays
                        .asList(PROPS.getProperty("targetWorkbookCols").split("\\|"));

        // target workbook col indices
        public static final int targetWorkbookAssetTagColIndex = Integer
                        .valueOf(PROPS.getProperty("targetWorkbookAssetTagColIndex")),
                        targetWorkbookSerialNumColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookSerialNumColIndex")),
                        targetWorkbookLocationColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookLocationColIndex")),
                        targetWorkbookRoomColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookRoomColIndex")),
                        targetWorkbookModelColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookModelColIndex")),
                        targetWorkbookUserColIndex = Integer.valueOf(PROPS.getProperty("targetWorkbookUserColIndex")),
                        targetWorkbookStatusColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookStatusColIndex")),
                        targetWorkbookBudgetNumColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookBudgetNumColIndex")),
                        targetWorkbookPurchDateColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookPurchDateColIndex")),
                        targetWorkbookPurchPriceColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookPurchPriceColIndex")),
                        targetWorkbookPurchOrderNumColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookPurchOrderNumColIndex")),
                        targetWorkbookLastInvDateColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookLastInvDateColIndex")),
                        targetWorkbookProductNumColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookProductNumColIndex")),
                        targetWorkbookModelNumColIndex = Integer
                                        .valueOf(PROPS.getProperty("targetWorkbookModelNumColIndex"));

        // messages and such
        public static final String cantFindAssetInInventoryMessage = PROPS
                        .getProperty("cantFindAssetInInventoryMessage"),
                        targetWorkbookAlreadyExistsMessage = PROPS.getProperty("targetWorkbookAlreadyExistsMessage"),
                        loadingAuditWorkbookMessage = PROPS.getProperty("loadingAuditWorkbookMessage"),
                        loadingInventoryWorkbookMessage = PROPS.getProperty("loadingInventoryWorkbookMessage"),
                        confirmationMessage = PROPS.getProperty("confirmationMessage"),
                        locationQuestionMessage = PROPS.getProperty("locationQuestionMessage"),
                        lastInvDateQuestionMessage = PROPS.getProperty("lastInvDateQuestionMessage"),
                        writingToTargetWorkbookMessage = PROPS.getProperty("writingToTargetWorkbookMessage"),
                        closingTargetWorkbookMessage = PROPS.getProperty("closingTargetWorkbookMessage");

        // console progress text messages
        public static final String createDevicesFromAuditWorkbookProgressMessage = PROPS
                        .getProperty("createDevicesFromAuditWorkbookProgressMessage"),
                        updateDevicesWithInventoryInfoProgressMessage = PROPS
                                        .getProperty("updateDevicesWithInventoryInfoProgressMessage"),
                        createTargetWorkbookWithDeviceInfoProgressMessage = PROPS
                                        .getProperty("createTargetWorkbookWithDeviceInfoProgressMessage");

        // Exception messages
        public static final String invalidDeviceStateMessage = PROPS.getProperty("invalidDeviceStateMessage"),
                        invalidRowStateMessage = PROPS.getProperty("invalidRowStateMessage");

        // misc
        public static final String decimalFormatPattern = PROPS.getProperty("decimalFormatPattern"),
                        targetWorkbookAlreadyExistsHandlerYesOption = PROPS
                                        .getProperty("targetWorkbookAlreadyExistsHandlerYesOption"),
                        targetWorkbookAlreadyExistsHandlerNoOption = PROPS
                                        .getProperty("targetWorkbookAlreadyExistsHandlerNoOption");

        // loads config.properties
        private static Properties loadProps(String filename) {
                try {
                        FileInputStream fis = new FileInputStream(filename);
                        Properties props = new Properties();
                        props.load(fis);
                        fis.close();
                        return props;
                } catch (IOException ioe) {
                        ioe.printStackTrace();
                        return null;
                }
        }

        private Config() {
                // can't be instantiated, since that's not the point of this class
        }
}
