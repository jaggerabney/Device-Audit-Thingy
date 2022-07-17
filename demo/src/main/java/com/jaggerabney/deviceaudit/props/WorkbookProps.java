package com.jaggerabney.deviceaudit.props;

public class WorkbookProps {
    private String name;

    public WorkbookProps() {
        this.name = null;
    }

    public WorkbookProps(String name, ColumnProps[] columns) {
        this.name = name;
    }

    public String name() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "Workbook " + name;
    }
}
