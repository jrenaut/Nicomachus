package docx;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Styles extends DocxXml {
    private boolean hasTableGrid = false;
    private List    lsdExceptions;
    private List    styles;

    public Styles() throws Exception {
        initXml();
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("w:styles");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        addLsdException("heading 2", 1);
    }

    protected Document writeXml() {
        int counter = 1;
        if (this.lsdExceptions != null) {
            Element latentStyles = getLatentStylesElement();
            for (Iterator it = this.lsdExceptions.iterator(); it.hasNext();) {
                LsdException r = (LsdException) it.next();
                latentStyles.appendChild(r.write(this.xml, counter++));
            }
            this.root.appendChild(latentStyles);
        }
        if (this.styles != null) {
            for (Iterator it = this.styles.iterator(); it.hasNext();) {
                Style s = (Style) it.next();
                this.root.appendChild(s.write(this.xml, counter++));
            }
        }
        this.xml.appendChild(this.root);
        return this.xml;
    }

    public void addLsdException(String name) {
        addLsdException(new LsdException(name));
    }

    public void addLsdException(String name, int qFormat) {
        addLsdException(new LsdException(name, -1, -1, qFormat));
    }

    public void addLsdException(String name, int semiHidden, int unhideWhenUsed, int qFormat) {
        addLsdException(new LsdException(name, semiHidden, unhideWhenUsed, qFormat));
    }

    public void addLsdException(String name, int semiHidden, int unhideWhenUsed) {
        addLsdException(new LsdException(name, semiHidden, unhideWhenUsed));
    }

    private void addLsdException(LsdException le) {
        if (this.lsdExceptions == null) {
            this.lsdExceptions = new ArrayList();
        }
        this.lsdExceptions.add(le);
    }

    public void addStyle(String id, String name, String basedOn, String type) {
        addStyle(new Style(id, name, basedOn, type));
    }

    /**
     * TODO - need to make sure style names are unique, don't want to add the
     * same style multiple times
     * 
     * @param style
     */
    public void addStyle(Style style) {
        if (this.styles == null)
            this.styles = new ArrayList();
        this.styles.add(style);
    }

    /**
     * TODO - this should be generic. Currently customized for Treasury 3.2.x
     * loop
     */
    public void addTableGrid() {
        if (!this.hasTableGrid) {
            Style s = new Style("TableOnePage", "Table One Page", "TableNormal", "table");
            s.addPprItem("w:keepNext");
            s.addPprItem("w:keepLines");
            s.setSpacing(0, 200);
            addStyle(s);
            this.hasTableGrid = true;
        }
    }

    public void addDefaultStyle() {
        Style s = new Style("Normal", "Normal", null, "paragraph");
        s.addQformat();
        addStyle(s);
    }

    public boolean hasTableGrid() {
        return this.hasTableGrid;
    }

    private Element getLatentStylesElement() {
        Element el = this.xml.createElement("w:latentStyles");
        el.setAttribute("w:count", "267");
        el.setAttribute("w:defLockedState", "0");
        el.setAttribute("w:defQFormat", "0");
        el.setAttribute("w:defSemiHidden", "1");
        el.setAttribute("w:defUIPriority", "99");
        el.setAttribute("w:defUnhideWhenUsed", "1");
        return el;
    }

    private class LsdException {
        private String name;
        private int    qFormat        = -1;
        private int    semiHidden     = -1;
        private int    unhideWhenUsed = -1;

        public LsdException(String name, int semiHidden, int unhideWhenUsed, int qFormat) {
            init(name, semiHidden, unhideWhenUsed, qFormat);
        }

        public LsdException(String name, int semiHidden, int unhideWhenUsed) {
            init(name, semiHidden, unhideWhenUsed, -1);
        }

        public LsdException(String name) {
            init(name, -1, -1, -1);
        }

        private void init(String name, int semiHidden, int unhideWhenUsed, int qFormat) {
            this.name = name;
            this.qFormat = qFormat;
            this.semiHidden = semiHidden;
            this.unhideWhenUsed = unhideWhenUsed;
        }

        public Element write(Document doc, int priority) {
            Element e = doc.createElement("w:lsdException");
            e.setAttribute("w:name", this.name);
            e.setAttribute("w:uiPriority", "" + priority);
            if (this.qFormat != -1)
                e.setAttribute("w:qFormat", "" + this.qFormat);
            if (this.semiHidden != -1)
                e.setAttribute("w:semiHidden", "" + this.semiHidden);
            if (this.unhideWhenUsed != -1)
                e.setAttribute("w:unhideWhenUsed", "" + this.unhideWhenUsed);
            return e;
        }
    }
}
