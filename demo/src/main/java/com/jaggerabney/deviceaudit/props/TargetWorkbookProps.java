package com.jaggerabney.deviceaudit.props;

public class TargetWorkbookProps extends WorkbookProps {
    private String sheetName;
    private NamedColumnProps[] columns;

    public TargetWorkbookProps() {
        super();
        this.sheetName = null;
        this.columns = null;
    }

    public TargetWorkbookProps(String name, NamedColumnProps[] columns, String sheetName) {
        super(name, columns);
        this.sheetName = sheetName;
        this.columns = columns;
    }

    public String sheetName() {
        return sheetName;
    }

    public NamedColumnProps[] columns() {
        return columns;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public void setColumns(NamedColumnProps[] columns) {
        this.columns = columns;
    }

    @Override
    public String toString() {
        String columnsStrings = "";

        for (NamedColumnProps column : columns) {
            columnsStrings += (column.toString() + "\n");
        }

        return super.toString() + " with sheet name " + sheetName + " and columns:\n" + columnsStrings;
    }
}
