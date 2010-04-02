package docx;

import java.util.ArrayList;
import java.util.Iterator;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class ContentTypes extends DocxXml {
    private ArrayList          defaults;
    private ArrayList          overrides;
    private boolean            hasInsertedDocx = false;
    private boolean            hasInsertedRtf  = false;
    private boolean            hasInsertedXls  = false;
    private boolean            hasInsertedDoc  = false;
    public static final String PATH            = "[Content_Types].xml";

    public ContentTypes() throws Exception {
        initXml();
    }

    public void addInsertedDocx() {
        if (!hasInsertedDocx) {
            hasInsertedDocx = true;
            addDefault("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
        }
    }

    public void addInsertedRtf() {
        if (!hasInsertedRtf) {
            hasInsertedRtf = true;
            addDefault("rtf", "application/rtf");
        }
    }

    public void addInsertedXls() {
        if (!hasInsertedXls) {
            hasInsertedXls = true;
            addDefault("xls", "application/ms-excel");
        }
    }

    public void addInsertedDoc() {
        if (!hasInsertedDoc) {
            hasInsertedDoc = true;
            addDefault("doc", "application/msword");
        }
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("Types");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");
        addDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
        addDefault("xml", "application/xml");
        addOverride("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
    }

    protected Document writeXml() {
        if (this.defaults != null) {
            for (Iterator it = this.defaults.iterator(); it.hasNext();) {
                Default d = (Default) it.next();
                this.root.appendChild(d.write(this.xml));
            }
        }
        if (this.overrides != null) {
            for (Iterator it = this.overrides.iterator(); it.hasNext();) {
                Override o = (Override) it.next();
                this.root.appendChild(o.write(this.xml));
            }
        }
        this.xml.appendChild(this.root);
        return this.xml;
    }

    public void addDefault(String extension, String contentType) {
        if (this.defaults == null) {
            this.defaults = new ArrayList();
        }
        Default d = new Default(extension, contentType);
        this.defaults.add(d);
    }

    public void addOverride(String partName, String contentType) {
        if (this.overrides == null) {
            this.overrides = new ArrayList();
        }
        Override o = new Override(partName, contentType);
        this.overrides.add(o);
    }

    private class Default {
        private final String ELEMENT_NAME      = "Default";
        private final String ATTR_EXTENSION    = "Extension";
        private final String ATTR_CONTENT_TYPE = "ContentType";
        private String       extension;
        private String       contentType;

        public Default(String ext, String type) {
            this.extension = ext;
            this.contentType = type;
        }

        public Element write(Document doc) {
            Element r = doc.createElement(ELEMENT_NAME);
            r.setAttribute(ATTR_CONTENT_TYPE, this.contentType);
            r.setAttribute(ATTR_EXTENSION, this.extension);
            return r;
        }
    }

    private class Override {
        private final String ELEMENT_NAME      = "Override";
        private final String ATTR_PARTNAME     = "PartName";
        private final String ATTR_CONTENT_TYPE = "ContentType";
        private String       partName;
        private String       contentType;

        public Override(String part, String type) {
            this.partName = part;
            this.contentType = type;
        }

        public Element write(Document doc) {
            Element r = doc.createElement(ELEMENT_NAME);
            r.setAttribute(ATTR_CONTENT_TYPE, this.contentType);
            r.setAttribute(ATTR_PARTNAME, this.partName);
            return r;
        }
    }
}
