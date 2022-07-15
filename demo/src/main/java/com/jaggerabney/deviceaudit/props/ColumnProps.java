package com.jaggerabney.deviceaudit.props;

public class ColumnProps {
    private String key;
    private int index;

    public ColumnProps() {
        this.key = null;
        this.index = 0;
    }

    public ColumnProps(String key, int index) {
        this.key = key;
        this.index = index;
    }

    public String key() {
        return key;
    }

    public int index() {
        return index;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public void setIndex(int index) {
        this.index = index;
    }
}
