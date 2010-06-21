package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;

public class FontTable extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("w:fonts");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
