package com.jaggerabney.deviceaudit;

// simple object used to track the pertinent information needed for a given row in target.xlsx
public final class Device {
    public String asset;
    public String serial;
    public String location;
    public String room;
    public String model;
    public String cot; // checked out to
    public String status;
    public String budgetNum;
    public String purchDate;
    public String purchPrice;
    public String purchOrderNum;
    public String lastInvDate;
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
    // for use in updateDevicesWithInventoryInfo
    public Device(String asset, String serial, String location, String room, String model, String cot, String status,
            String budgetNum, String purchDate, String purchPrice, String purchOrderNum, String lastInvDate,
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
        this.lastInvDate = lastInvDate;
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
