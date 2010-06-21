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

public class RelsRels extends DocxXml {
    private ArrayList          relationships;
    public static final String PATH = "_rels/.rels";

    public RelsRels() throws Exception {
        initXml();
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("Relationships");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");
        addRelationship("word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    }

    public Document writeXml() {
        if (this.relationships == null) {
            return this.xml;
        }
        for (Iterator it = this.relationships.iterator(); it.hasNext();) {
            Relationship r = (Relationship) it.next();
            this.root.appendChild(r.write(this.xml));
        }
        this.xml.appendChild(this.root);
        return this.xml;
    }

    public void addRelationship(String target, String type) {
        if (this.relationships == null) {
            this.relationships = new ArrayList();
        }
        Relationship r = new Relationship(target, type, this.relationships.size() + 1);
        this.relationships.add(r);
    }

    /**
     * @param args
     */
    public static void main(String[] args) {
        try {
            RelsRels xmlrels = new RelsRels();
            xmlrels.addRelationship("tag1", "type1");
            xmlrels.addRelationship("tag2", "type2");
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
        private final String ELEMENT_NAME = "Relationship";
        private final String ATTR_TARGET  = "Target";
        private final String ATTR_TYPE    = "Type";
        private final String ATTR_ID      = "Id";
        private String       target;
        private String       type;
        private int          id;

        public Relationship(String target, String type, int id) {
            this.target = target;
            this.type = type;
            this.id = id;
        }

        public Element write(Document doc) {
            Element r = doc.createElement(ELEMENT_NAME);
            r.setAttribute(ATTR_TARGET, this.target);
            r.setAttribute(ATTR_TYPE, this.type);
            r.setAttribute(ATTR_ID, "rId" + this.id);
            return r;
        }
    }
}
