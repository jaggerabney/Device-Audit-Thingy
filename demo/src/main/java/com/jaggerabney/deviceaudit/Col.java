package com.jaggerabney.deviceaudit;

public final class Col {
    public String name;
    public int index;

    private Col(String name, int index) {
        this.name = name;
        this.index = index;
    }

    public static Col newCol(String name, int index) {
        return new Col(name, index);
    }

    @Override
    public String toString() {
        return "(" + name + ", " + index + ")";
    }
}
