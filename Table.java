package docx;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

public class Table {
    private List     rows;
    private double[] columnWidths;
    private double   width = 0;

    public void addRow(Row r) {
        if (rows == null) {
            rows = new ArrayList();
        }
        rows.add(r);
    }

    public void addBlankRow(int columnCount) {
        addBlankRow(columnCount, Color.WHITE);
    }

    public void addBlankRow(int columnCount, Color bg) {
        Row r = new Row();
        for (int i = 0; i < columnCount; i++) {
            r.addCell(new Cell("", bg, bg));
        }
        addRow(r);
    }

    public List getRows() {
        return this.rows;
    }

    public void setColumnWidths(double[] cw) {
        this.columnWidths = new double[cw.length];
        for (int i = 0; i < cw.length; i++) {
            this.columnWidths[i] = cw[i] * 1440;
        }
    }

    public double[] getColumnWidths() {
        return this.columnWidths;
    }

    public void setWidth(double inches) {
        this.width = inches;
    }

    public double getWidth() {
        return (1440 * this.width);
    }

    public static void main(String[] args) {
        // TODO Auto-generated method stub
    }
}
