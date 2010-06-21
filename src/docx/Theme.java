package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Theme extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        Element root = this.xml.createElement("a:theme");
        root.setAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        root.setAttribute("name", "Office Theme");
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
