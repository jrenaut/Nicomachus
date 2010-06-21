package docx;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Style {
    private String    id;
    private String    name;
    private String    basedOn;
    private String    type;
    private ArrayList pprItems;
    private ArrayList childItems;
    private ArrayList rprItems;
    private boolean   qFormat = false;

    public Style(String id, String name, String basedOn, String type) {
        this.id = id;
        this.name = name;
        this.basedOn = basedOn;
        this.type = type;
    }

    public Style(String id, String type) {
        this.id = id;
        this.type = type;
    }

    public void addPprItem(String itemName) {
        addPprItem(new ChildItem(itemName));
    }

    public void addQformat() {
        this.qFormat = true;
    }

    public void addPprItem(ChildItem p) {
        if (this.pprItems == null)
            this.pprItems = new ArrayList();
        this.pprItems.add(p);
    }

    public void addChildItem(ChildItem c) {
        if (this.childItems == null)
            this.childItems = new ArrayList();
        childItems.add(c);
    }

    public void addRprItem(ChildItem c) {
        if (this.rprItems == null)
            this.rprItems = new ArrayList();
        rprItems.add(c);
    }

    public void setFont(String fontName) {
        ChildItem p = new ChildItem("w:rFonts");
        p.setAttribute("w:ascii", fontName);
        p.setAttribute("w:hAnsi", fontName);
        p.setAttribute("w:cs", fontName);
        addRprItem(p);
    }

    public void setColor(Color color) {
        ChildItem p = new ChildItem("w:color");
        p.setAttribute("w:val", convertColorToHexString(color));
        addRprItem(p);
    }

    public void setFontSize(int size) {
        ChildItem p = new ChildItem("w:sz");
        p.setAttribute("w:val", "" + (2 * size));
        addRprItem(p);
        p = new ChildItem("w:szCs");
        p.setAttribute("w:val", "" + (2 * size));
        addRprItem(p);
    }

    public void setBold() {
        addRprItem(new ChildItem("w:b"));
        addRprItem(new ChildItem("w:bCs"));
    }

    public void setSpacing(int before, int after) {
        ChildItem p = new ChildItem("w:spacing");
        p.setAttribute("w:after", "" + after);
        p.setAttribute("w:before", "" + before);
        addPprItem(p);
    }

    public void setOutlineLevel(int size) {
        ChildItem p = new ChildItem("w:outlineLvl");
        p.setAttribute("w:val", "" + size);
        addPprItem(p);
    }

    public void setUnhideWhenUsed() {
        addChildItem(new ChildItem("w:unhideWhenUsed"));
    }

    public void setQFormat() {
        addChildItem(new ChildItem("w:qFormat"));
    }

    public void setKeepNext() {
        addPprItem(new ChildItem("w:keepNext"));
    }

    public void setKeepLines() {
        addPprItem(new ChildItem("w:keepLines"));
    }

    public void setUiPriority(int priority) {
        ChildItem p = new ChildItem("w:uiPriority");
        p.setAttribute("w:val", "" + priority);
        addChildItem(p);
    }

    public void setLink(String link) {
        ChildItem p = new ChildItem("w:link");
        p.setAttribute("w:val", link);
        addChildItem(p);
    }

    public void setNext(String val) {
        ChildItem p = new ChildItem("w:next");
        p.setAttribute("w:val", val);
        addChildItem(p);
    }

    public Element write(Document doc, int priority) {
        Element e;
        Element style = doc.createElement("w:style");
        style.setAttribute("w:styleId", id);
        style.setAttribute("w:type", type);
        if (this.name != null) {
            e = doc.createElement("w:name");
            e.setAttribute("w:val", name);
            style.appendChild(e);
        }
        if (this.basedOn != null) {
            e = doc.createElement("w:basedOn");
            e.setAttribute("w:val", basedOn);
            style.appendChild(e);
        }
        e = doc.createElement("w:uiPriority");
        e.setAttribute("w:val", "" + priority);
        style.appendChild(e);
        if (this.qFormat) {
            e = doc.createElement("w:qFormat");
            style.appendChild(e);
        }
        if (this.pprItems != null) {
            Element ppr = doc.createElement("w:pPr");
            for (Iterator it = this.pprItems.iterator(); it.hasNext();) {
                ChildItem item = (ChildItem) it.next();
                ppr.appendChild(item.write(doc));
            }
            style.appendChild(ppr);
        }
        if (this.rprItems != null) {
            Element rpr = doc.createElement("w:rPr");
            for (Iterator it = this.rprItems.iterator(); it.hasNext();) {
                ChildItem item = (ChildItem) it.next();
                rpr.appendChild(item.write(doc));
            }
            style.appendChild(rpr);
        }
        if (this.childItems != null) {
            for (Iterator it = this.childItems.iterator(); it.hasNext();) {
                ChildItem item = (ChildItem) it.next();
                style.appendChild(item.write(doc));
            }
        }
        return style;
    }

    /**
     * Converts a java.awt.Color object to its hex equivalent TODO - copied from
     * Docx.java. Decide where this makes sense and put it there.
     * 
     * @param color
     * @return
     */
    private String convertColorToHexString(Color color) {
        String retval = Integer.toHexString(color.getRGB() & 0x00ffffff);
        if (retval.length() == 5) {
            // TODO - find a better way to keep it from dropping leading 0
            retval = "0" + retval;
        }
        return retval;
    }

    private class ChildItem {
        private String name;
        private Map    attributes;

        public ChildItem(String name) {
            this.name = name;
        }

        public void setAttribute(String k, String v) {
            if (this.attributes == null)
                this.attributes = new HashMap();
            this.attributes.put(k, v);
        }

        public Element write(Document doc) {
            Element ppr = doc.createElement(this.name);
            if (this.attributes != null) {
                for (Iterator it = this.attributes.keySet().iterator(); it.hasNext();) {
                    String k = (String) it.next();
                    String v = (String) this.attributes.get(k);
                    ppr.setAttribute(k, v);
                }
            }
            return ppr;
        }
    }
}
