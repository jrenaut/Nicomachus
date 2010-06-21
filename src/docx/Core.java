package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;

public class Core extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("cp:coreProperties");
        this.root.setAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        this.root.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
        this.root.setAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
        this.root.setAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
        this.root.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
