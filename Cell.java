package docx;

import java.awt.Color;
import java.util.ArrayList;

public class Cell {
    private boolean   bold            = false;
    private boolean   italic          = false;
    private String    text;
    private Color     backgroundColor = Color.WHITE;
    private Color     fontColor       = Color.BLACK;
    private double    width           = -1;
    private int       rowSpan         = 1;
    private int       colSpan         = 1;
    private ArrayList borders;
    private int       alignment;
    private final int ALIGN_LEFT      = 0;
    private final int ALIGN_RIGHT     = 1;
    private final int ALIGN_CENTER    = 2;
    private int       fontSize        = 11;
    private String    fontName        = "Times New Roman";

    public int getColumnSpan() {
        return this.colSpan;
    }

    public void setColumnSpan(int colSpan) {
        this.colSpan = colSpan;
    }

    public int getRowSpan() {
        return this.rowSpan;
    }

    public void setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
        throw new RuntimeException("Too complicated to implement now");
    }

    public Cell(String text) {
        this.text = text;
    }

    public Cell(String text, Color fontColor, Color backgroundColor) {
        this.text = text;
        this.fontColor = fontColor;
        this.backgroundColor = backgroundColor;
    }

    public void setBold() {
        this.bold = true;
    }

    public void setItalic() {
        this.italic = true;
    }

    public void setBackgroundColor(Color c) {
        this.backgroundColor = c;
    }

    public void setFontColor(Color c) {
        this.fontColor = c;
    }

    public boolean isBold() {
        return this.bold;
    }

    public boolean isItalic() {
        return this.italic;
    }

    public String getText() {
        return this.text;
    }

    public Color getBackgroundColor() {
        return this.backgroundColor;
    }

    public Color getFontColor() {
        return this.fontColor;
    }

    public double getWidth() {
        return this.width;
    }

    public void setWidth(double widthInInches) {
        this.width = 1440 * widthInInches;
    }

    public void addBorder(CellBorder border) {
        if (this.borders == null) {
            borders = new ArrayList();
        }
        borders.add(border);
    }

    public ArrayList getBorders() {
        return this.borders;
    }

    public void alignLeft() {
        this.alignment = ALIGN_LEFT;
    }

    public void alignRight() {
        this.alignment = ALIGN_RIGHT;
    }

    public void alignCenter() {
        this.alignment = ALIGN_CENTER;
    }

    public String getAlignment() {
        switch (this.alignment) {
        case ALIGN_RIGHT:
            return "right";
        case ALIGN_CENTER:
            return "center";
        case ALIGN_LEFT:
            return "left";
        default:
            return "left";
        }
    }

    public void setFont(String fontName, int fontSize) {
        setFontName(fontName);
        setFontSize(fontSize);
    }

    public String getFontName() {
        return this.fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public int getFontSize() {
        return this.fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }
}
