package com.jaggerabney.deviceaudit;

import java.util.*;

// simple object used to track the pertinent information needed for a given row in target.xlsx
public final class Device {
    public static final int NUM_PROPS = 14; // update as you add props!!!

    public String asset;
    public String serial;
    public String location;
    public String room;
    public String model;
    public String cot; // checked out to
    public String status;
    public String budgetNum;
    public String purchDate;
    public double purchPrice;
    public String purchOrderNum;
    public Date lastInventoryDate;
    public String productNum;
    public String modelNum;

    // for use in createDevicesFromWorkbook
    public Device(String room, String assetOrSerial, String cot, boolean isAssetTag) {
        this.room = room;
        if (isAssetTag) {
            this.asset = assetOrSerial;
        } else {
            this.serial = assetOrSerial;
        }
        this.cot = cot;
    }

    // lol. lmao
    // for use in updateDevicesWithInventoryInfo
    public Device(String asset, String serial, String location, String room, String model, String cot,
            String status) {
        this.asset = asset;
        this.serial = serial;
        this.location = location;
        this.room = room;
        this.model = model;
        this.cot = cot;
        this.status = status;
    }

    // oh hell
    public Device(String asset, String serial, String location, String room, String model, String cot, String status,
            String budgetNum, String purchDate, double purchPrice, String purchOrderNum, Date lastInventoryDate,
            String productNum, String modelNum) {
        this.asset = asset;
        this.serial = serial;
        this.location = location;
        this.room = room;
        this.model = model;
        this.cot = cot;
        this.status = status;
        this.budgetNum = budgetNum;
        this.purchDate = purchDate;
        this.purchPrice = purchPrice;
        this.purchOrderNum = purchOrderNum;
        this.lastInventoryDate = lastInventoryDate;
        this.productNum = productNum;
        this.modelNum = modelNum;
    }

    // for debugging and the like
    @Override
    public String toString() {
        return "[" + asset + ", " + serial + ", " + location + ", " + room + ", " + model + ", " + cot + ", " + status
                + "]\n";
    }
}
