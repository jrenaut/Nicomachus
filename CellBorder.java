package docx;

import java.awt.Color;

public class CellBorder {
    public static final int BORDER_LEFT   = 0;
    public static final int BORDER_BOTTOM = 1;
    public static final int BORDER_RIGHT  = 2;
    public static final int BORDER_TOP    = 3;
    public static final int BORDER_SINGLE = 0;
    private int             location;
    private int             size;
    private int             type;
    private Color           color;

    public CellBorder(int location, int size, int type, Color color) {
        init(location, size, type, color);
    }

    public CellBorder(int location) {
        init(location, 12, BORDER_SINGLE, Color.BLACK);
    }

    private void init(int location, int size, int type, Color color) {
        setLocation(location);
        setSize(size);
        setType(type);
        setColor(color);
    }

    public Color getColor() {
        return this.color;
    }

    public void setColor(Color color) {
        this.color = color;
    }

    public String getLocation() {
        switch (this.location) {
        case BORDER_BOTTOM:
            return "bottom";
        case BORDER_LEFT:
            return "left";
        case BORDER_RIGHT:
            return "right";
        case BORDER_TOP:
            return "top";
        default:
            return "bottom";
        }
    }

    public void setLocation(int location) {
        this.location = location;
    }

    public int getSize() {
        return this.size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    public String getType() {
        switch (this.type) {
        case BORDER_SINGLE:
            return "single";
        default:
            return "single";
        }
    }

    public void setType(int type) {
        this.type = type;
    }
}
