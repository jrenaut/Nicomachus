package docx;

import java.util.ArrayList;
import java.util.List;

public class Row {
    private List cells;

    public void addCell(Cell c) {
        if (cells == null) {
            cells = new ArrayList();
        }
        cells.add(c);
    }

    public List getCells() {
        return this.cells;
    }
}
