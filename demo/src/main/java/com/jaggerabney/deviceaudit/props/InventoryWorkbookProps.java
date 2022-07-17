package com.jaggerabney.deviceaudit.props;

public class InventoryWorkbookProps extends WorkbookProps {
    private NamedColumnProps[] columns;

    public InventoryWorkbookProps() {
        super();
        this.columns = null;
    }

    public InventoryWorkbookProps(NamedColumnProps[] columns) {
        this.columns = columns;
    }

    public NamedColumnProps[] columns() {
        return columns;
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

        return " and columns:\n" + columnsStrings;
    }
}
