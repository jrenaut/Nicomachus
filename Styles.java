package docx;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Styles extends DocxXml {
    public Styles() throws Exception {
        initXml();
    }

    protected void initXml() throws Exception {
        this.xml = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        this.root = this.xml.createElement("w:styles");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }

    protected Document writeXml() {
        this.root.appendChild(addDefaults());
        this.xml.appendChild(this.root);
        return this.xml;
    }

    private Element addDefaults() {
        Element docDefaults = this.xml.createElement("w:docDefaults");
        Element rPrDefault = this.xml.createElement("w:rPrDefault");
        Element rpr = this.xml.createElement("w:rPr");
        Element fonts = this.xml.createElement("w:rFonts");
        fonts.setAttribute("w:ascii", "Calibri");
        Element sz = this.xml.createElement("w:sz");
        sz.setAttribute("w:val", "24");
        rpr.appendChild(fonts);
        rpr.appendChild(sz);
        rPrDefault.appendChild(rpr);
        Element pPrDefault = this.xml.createElement("wpPrDefault");
        Element pPr = this.xml.createElement("w:pPr");
        Element spacing = this.xml.createElement("w:spacing");
        spacing.setAttribute("w:after", "120");
        pPr.appendChild(spacing);
        pPrDefault.appendChild(pPr);
        docDefaults.appendChild(rPrDefault);
        docDefaults.appendChild(pPrDefault);

        return docDefaults;
    }

}
