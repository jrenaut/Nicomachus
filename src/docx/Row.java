package docx;

import java.util.ArrayList;
import java.util.List;

public class Row {
    private List    cells;
    private boolean bap       = false;
    private double  rowHeight = 0;
    private boolean addTrPr   = false;

    public void addCell(Cell c) {
        if (cells == null) {
            cells = new ArrayList();
        }
        cells.add(c);
    }

    public List getCells() {
        return this.cells;
    }

    public boolean canBreakAcrossPages() {
        return this.bap;
    }

    /**
     * Allow row to break across pages
     * 
     * @param canBreak
     */
    public void setCanBreakAcrossPages(boolean canBreak) {
        this.bap = canBreak;
        this.addTrPr = true;
    }

    /**
     * Setting height at row level will override table level setting
     * 
     * @param inches
     */
    public void setRowHeight(double inches) {
        this.rowHeight = inches;
        this.addTrPr = true;
    }

    /**
     * Return row height as required by docx (inches times 1440)
     * 
     * @return
     */
    public double getRowHeight() {
        return (1440 * this.rowHeight);
    }

    /**
     * True if row has attributes set so that w:trPr element is necessary
     * 
     * @return
     */
    public boolean needsTrPr() {
        return this.addTrPr;
    }
}
