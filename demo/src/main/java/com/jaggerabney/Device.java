package com.jaggerabney;

public class Device {
    // look i know making fields public is a bad idea but i dont want to use a
    // million getters okay
    public String asset;
    public String serial;
    public String model;
    public String status;

    public Device(String asset) {
        this.asset = asset;
    }
}
