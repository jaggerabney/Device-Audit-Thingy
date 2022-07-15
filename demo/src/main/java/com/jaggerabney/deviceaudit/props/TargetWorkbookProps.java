package com.jaggerabney.deviceaudit.props;

public class TargetWorkbookProps extends WorkbookProps {
    private String sheetName;

    public TargetWorkbookProps() {
        super();
        this.sheetName = null;
    }

    public TargetWorkbookProps(String name, ColumnProps[] columns, String sheetName) {
        super(name, columns);
        this.sheetName = sheetName;
    }

    public String sheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}
