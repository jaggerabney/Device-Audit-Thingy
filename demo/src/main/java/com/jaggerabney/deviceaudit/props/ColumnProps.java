package com.jaggerabney.deviceaudit.props;

import org.apache.poi.ss.formula.functions.Column;

public final class ColumnProps {
    public String key;
    public int index;

    private ColumnProps(String key, int index) {
        this.key = null;
        this.index = 0;
    }

    public static ColumnProps newColumnsProps(String key, int index) {
        return new ColumnProps(key, index);
    }
}
