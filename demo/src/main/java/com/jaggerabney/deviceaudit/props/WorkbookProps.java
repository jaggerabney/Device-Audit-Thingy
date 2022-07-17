package com.jaggerabney.deviceaudit.props;

public class WorkbookProps {
    private String name;
    private ColumnProps[] columns;

    public WorkbookProps() {
        this.name = null;
        this.columns = null;
    }

    public WorkbookProps(String name, ColumnProps[] columns) {
        this.name = name;
        this.columns = columns;
    }

    public String name() {
        return name;
    }

    public ColumnProps[] columns() {
        return columns;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setColumns(ColumnProps[] columns) {
        this.columns = columns;
    }

    @Override
    public String toString() {
        String columnsStrings = "";

        for (ColumnProps column : columns) {
            columnsStrings += column.toString();
            columnsStrings += "\n";
        }

        return "Workbook " + name + " with columns: " + columnsStrings;
    }
}
