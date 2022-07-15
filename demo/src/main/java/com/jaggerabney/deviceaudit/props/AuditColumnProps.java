package com.jaggerabney.deviceaudit.props;

public class AuditColumnProps extends ColumnProps {
    private boolean allowEmptySpaces;

    public AuditColumnProps() {
        super();
        this.allowEmptySpaces = false;
    }

    public AuditColumnProps(String key, int index, boolean allowEmptySpaces) {
        super(key, index);
        this.allowEmptySpaces = allowEmptySpaces;
    }

    public boolean allowEmptySpaces() {
        return allowEmptySpaces;
    }

    public void setAllowEmptySpaces(boolean allowEmptySpaces) {
        this.allowEmptySpaces = allowEmptySpaces;
    }
}
