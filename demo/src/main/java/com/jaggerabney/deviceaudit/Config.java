package com.jaggerabney.deviceaudit;

import java.util.*;
import java.io.*;

public class Config {
    private static final Properties PROPS = loadProps("config.properties");

    public static final int auditWorkbookRoomColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookRoomColIndex")),
            auditWorkbookAssetColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookAssetColIndex")),
            auditWorkbookUserColIndex = Integer.valueOf(PROPS.getProperty("auditWorkbookUserColIndex")),
            auditWorkbookAssetThreshold = Integer.valueOf(PROPS.getProperty("auditWorkbookAssetThreshold"));

    public static final String auditWorkbookName = PROPS.getProperty("auditWorkbookName"),
            inventoryWorkbookName = PROPS.getProperty("inventoryWorkbookName"),
            targetWorkbookName = PROPS.getProperty("targetWorkbookName"),
            targetWorkbookSheetName = PROPS.getProperty("targetWorkbookSheetName");

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

    public static final List<String> targetWorkbookColumns = Arrays
            .asList(PROPS.getProperty("targetWorkbookCols").split("|"));

    public static final String cantFindAssetInInventoryMessage = PROPS.getProperty("cantFindAssetInInventoryMessage"),
            targetWorkbookAlreadyExistsMessage = PROPS.getProperty("targetWorkbookAlreadyExistsMessage"),
            loadingAuditWorkbookMessage = PROPS.getProperty("loadingAuditWorkbookMessage"),
            loadingInventoryWorkbookMessage = PROPS.getProperty("loadingInventoryWorkbookMessage"),
            confirmationMessage = PROPS.getProperty("confirmationMessage"),
            locationQuestionMessage = PROPS.getProperty("locationQuestionMessage"),
            writingToTargetWorkbookMessage = PROPS.getProperty("writingToTargetWorkbookMessage"),
            closingTargetWorkbookMessage = PROPS.getProperty("closingTargetWorkbookMessage");

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
