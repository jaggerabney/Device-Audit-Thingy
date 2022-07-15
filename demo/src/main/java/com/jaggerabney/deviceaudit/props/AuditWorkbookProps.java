package com.jaggerabney.deviceaudit.props;

public class AuditWorkbookProps extends WorkbookProps {
    private int assetLengthThreshold;

    public AuditWorkbookProps() {
        super();
        this.assetLengthThreshold = 0;
    }

    public AuditWorkbookProps(String name, ColumnProps[] columns, int assetLengthThreshold) {
        super(name, columns);
        this.assetLengthThreshold = assetLengthThreshold;
    }

    public int assetLengthThreshold() {
        return assetLengthThreshold;
    }

    public void setAssetLengthThreshold(int assetLengthThreshold) {
        this.assetLengthThreshold = assetLengthThreshold;
    }
}
