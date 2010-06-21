package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Settings extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        Element root = this.xml.createElement("w:settings");
        root.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
        root.setAttribute("xmlns:v", "urn:schemas-microsoft-com:vml");
        root.setAttribute("xmlns:w10", "urn:schemas-microsoft-com:office:word");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        root.setAttribute("xmlns:sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
