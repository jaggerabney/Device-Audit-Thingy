package com.jaggerabney.deviceaudit.props;

public class NamedColumnProps extends ColumnProps {
    private String name;

    public NamedColumnProps() {
        super();
        this.name = null;
    }

    public NamedColumnProps(String key, int index, String name) {
        super(key, index);
        this.name = name;
    }

    public String name() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
