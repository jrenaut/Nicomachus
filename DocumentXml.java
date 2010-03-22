package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class DocumentXml extends DocxXml {
    private Element body;

    public DocumentXml() throws Exception {
        initXml();
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("w:document");
        root.setAttribute("xmlns:ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        root.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
        root.setAttribute("xmlns:v", "urn:schemas-microsoft-com:vml");
        root.setAttribute("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        root.setAttribute("xmlns:w10", "urn:schemas-microsoft-com:office:word");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        root.setAttribute("xmlns:wne", "http://schemas.microsoft.com/office/word/2006/wordml");
        this.root.appendChild(getBody());
    }

    public Element getBody() {
        if (this.body == null) {
            this.body = this.xml.createElement("w:body");
        }
        return this.body;
    }

    public Document getDocument() {
        return this.xml;
    }

    public void setPreserveSpace(boolean preserveSpace) {
        if (preserveSpace) {
            this.body.setAttribute("xml:space", "preserve");
        }
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }

    /**
     * @param args
     */
    public static void main(String[] args) {
        // TODO Auto-generated method stub
    }
}
