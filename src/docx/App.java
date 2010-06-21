package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;

public class App extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("Properties");
        this.root.setAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        this.root.setAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
