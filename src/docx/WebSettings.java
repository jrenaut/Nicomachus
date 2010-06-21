package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class WebSettings extends DocxXml {
    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        Element root = this.xml.createElement("w:webSettings");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        Element el;
        el = this.xml.createElement("w:optimizeForBrowser");
        root.appendChild(el);
    }

    protected Document writeXml() {
        this.xml.appendChild(this.root);
        return this.xml;
    }
}
