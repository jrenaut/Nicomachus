package docx;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class DocumentXmlRels extends DocxXml {
    private ArrayList          relationships;
    public static final String PATH = "word/_rels/document.xml.rels";

    public DocumentXmlRels() throws Exception {
        initXml();
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("Relationships");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
    }

    public Document writeXml() {
        if (this.relationships == null) {
            return this.xml;
        }
        int counter = 1;
        for (Iterator it = this.relationships.iterator(); it.hasNext();) {
            Object o = it.next();
            if (o instanceof Relationship) {
                Relationship r = (Relationship) o;
                if (r.id.equals("NONE")) {
                    r.id = "rId" + counter;
                    counter++;
                }
                this.root.appendChild(r.write(this.xml));
            } else {
                SubDocument r = (SubDocument) o;
                if (r.id.equals("NONE")) {
                    r.id = "rId" + counter;
                    counter++;
                }
                this.root.appendChild(r.write(this.xml));
            }
        }
        this.xml.appendChild(this.root);
        return this.xml;
    }

    public void addRelationship(String target, String type, String id) {
        if (this.relationships == null) {
            this.relationships = new ArrayList();
        }
        Relationship r = new Relationship(target, type, id);
        this.relationships.add(r);
    }

    public void addRelationship(String target, String type) {
        if (this.relationships == null) {
            this.relationships = new ArrayList();
        }
        Relationship r = new Relationship(target, type);
        this.relationships.add(r);
    }

    public void addSubDocument(String target, String type, String id) {
        if (this.relationships == null) {
            this.relationships = new ArrayList();
        }
        SubDocument r = new SubDocument(target, type, id);
        this.relationships.add(r);
    }

    /**
     * @param args
     */
    public static void main(String[] args) {
        try {
            DocumentXmlRels xmlrels = new DocumentXmlRels();
            xmlrels.addRelationship("tag1", "type1", "x");
            xmlrels.addRelationship("tag2", "type2", "y");
            FileOutputStream fos = new FileOutputStream(new File("c:\\temp\\xmlrels.xml"));
            Transformer t = TransformerFactory.newInstance().newTransformer();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            t.transform(new DOMSource(xmlrels.writeXml()), new StreamResult(baos));
            byte[] data = baos.toByteArray();
            fos.write(data);
            fos.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private class Relationship {
        protected final String ELEMENT_NAME = "Relationship";
        protected final String ATTR_TARGET  = "Target";
        protected final String ATTR_TYPE    = "Type";
        protected final String ATTR_ID      = "Id";
        protected String       target;
        protected String       type;
        protected String       id           = "NONE";

        public Relationship(String target, String type, String id) {
            this.target = target;
            this.type = type;
            this.id = id;
        }

        public Relationship(String target, String type) {
            this.target = target;
            this.type = type;
        }

        public Element write(Document doc) {
            Element r = doc.createElement(ELEMENT_NAME);
            r.setAttribute(ATTR_TARGET, this.target);
            r.setAttribute(ATTR_TYPE, this.type);
            r.setAttribute(ATTR_ID, this.id);
            return r;
        }
    }

    private class SubDocument extends Relationship {
        private final String TARGET_MODE = "TargetMode";

        public SubDocument(String target, String type, String id) {
            super(target, type, id);
        }

        public Element write(Document doc) {
            Element r = doc.createElement(ELEMENT_NAME);
            r.setAttribute(ATTR_TARGET, this.target);
            r.setAttribute(ATTR_TYPE, this.type);
            r.setAttribute(ATTR_ID, this.id);
            r.setAttribute(TARGET_MODE, "External");
            return r;
        }
    }
}
