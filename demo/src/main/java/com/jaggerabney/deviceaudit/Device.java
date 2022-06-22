package com.jaggerabney.deviceaudit;

public final class Device {
    public int row;
    public String asset;
    public String serial;
    public String location;
    public String room;
    public String model;
    public String cot; // checked out to
    public String status;

    // for use in createDevicesFromWorkbook
    public Device(int row, String room, String asset, String cot) {
        this.row = row;
        this.room = room;
        this.asset = asset;
        this.cot = cot;
    }

    // lol. lmao
    // for use in <insert function name here>
    public Device(int row, String asset, String serial, String location, String room, String model, String cot,
            String status) {
        this.row = row;
        this.asset = asset;
        this.serial = serial;
        this.location = location;
        this.room = room;
        this.model = model;
        this.cot = cot;
        this.status = status;
    }

    @Override
    public String toString() {
        return "[" + asset + ", " + serial + ", " + location + ", " + room + ", " + model + ", " + cot + ", " + status
                + "]\n";
    }
}
