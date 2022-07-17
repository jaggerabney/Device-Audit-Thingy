package com.jaggerabney.deviceaudit.props;

public class AuditWorkbookProps extends WorkbookProps {
    private int assetLengthThreshold;
    private AuditColumnProps[] columns;

    public AuditWorkbookProps() {
        super();
        this.assetLengthThreshold = 0;
        this.columns = null;
    }

    public AuditWorkbookProps(String name, int assetLengthThreshold, AuditColumnProps[] columns) {
        super(name, columns);
        this.assetLengthThreshold = assetLengthThreshold;
        this.columns = columns;
    }

    public int assetLengthThreshold() {
        return assetLengthThreshold;
    }

    public AuditColumnProps[] columns() {
        return columns;
    }

    public void setAssetLengthThreshold(int assetLengthThreshold) {
        this.assetLengthThreshold = assetLengthThreshold;
    }

    public void setColumns(AuditColumnProps[] columns) {
        this.columns = columns;
    }

    @Override
    public String toString() {
        String columnsStrings = "";

        for (AuditColumnProps column : columns) {
            columnsStrings += (column.toString() + "\n");
        }

        return super.toString() + " with asset length threshold " + assetLengthThreshold + " and columns:\n"
                + columnsStrings;
    }
}
