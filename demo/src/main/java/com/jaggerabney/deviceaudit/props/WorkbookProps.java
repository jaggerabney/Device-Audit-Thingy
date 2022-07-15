package com.jaggerabney.deviceaudit.props;

public final class WorkbookProps {
    public String name;
    public ColumnProps[] columns;

    private WorkbookProps(String name, ColumnProps[] columns) {
        this.name = name;
        this.columns = columns;
    }

    public static WorkbookProps newWorkbookProps(String name, ColumnProps[] columns) {
        return new WorkbookProps(name, columns);
    }
}
