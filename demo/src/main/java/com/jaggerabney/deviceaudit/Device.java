package com.jaggerabney.deviceaudit;

// simple object used to track the pertinent information needed for a given row in target.xlsx
public final class Device {
    public static final int NUM_PROPS = 7; // update as you add props!!!

    public String asset;
    public String serial;
    public String location;
    public String room;
    public String model;
    public String cot; // checked out to
    public String status;

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

    // for debugging and the like
    @Override
    public String toString() {
        return "[" + asset + ", " + serial + ", " + location + ", " + room + ", " + model + ", " + cot + ", " + status
                + "]\n";
    }
}
