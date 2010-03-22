package docx;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import oracle.sql.BLOB;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

public class Docx {
    private DocumentXml            document_xml;
    private Document               fontTable_xml;
    private Document               settings_xml;
    private Styles                 styles_xml;
    private Document               webSettings_xml;
    private ContentTypes           ContentTypes_xml;
    private DocumentXmlRels        document_xml_rels;
    private Document               theme_xml;
    private RelsRels               rels_xml;
    private Document               app_xml;
    private Document               core_xml;
    private DocumentBuilderFactory dbf              = null;
    // NB - if you set ADD_OPTIONAL_XML to true, you get an invalid docx file.
    // Still working
    // on that. -- JER
    private boolean                ADD_OPTIONAL_XML = false;
    private ZipOutputStream        zos;
    private BLOB                   tempBlob;
    private boolean                isDebug          = false;

    public Docx() {
        try {
            dbf = DocumentBuilderFactory.newInstance();
            Connection conn = Utils.openConnection();
            this.tempBlob = BLOB.createTemporary(conn, false, BLOB.DURATION_SESSION);
            this.tempBlob.open(BLOB.MODE_READWRITE);
            this.tempBlob.truncate(0);
            this.zos = new ZipOutputStream(tempBlob.setBinaryStream(1L));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public Docx(String debugFileName) {
        this.isDebug = true;
        try {
            dbf = DocumentBuilderFactory.newInstance();
            this.zos = new ZipOutputStream(new FileOutputStream(debugFileName));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void appendChild(Element e) {
        getDocumentBody().appendChild(e);
    }

    private Element getDocumentBody() {
        try {
            return this.document_xml.getBody();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public void addInsertedDocument(BLOB b, String filename, String fileType, String fileId) throws Exception {
        Connection conn = Utils.openConnection();
        BLOB temp = Utils.createTemporaryBlob(conn);
        Utils.unzipBlobtoBlob("", b, temp);
        filename = "word/" + filename;
        ZipEntry ze = new ZipEntry(filename);
        InputStream inStream = temp.binaryStreamValue();
        int length = -1;
        int size = temp.getBufferSize();
        ze.setSize(size);
        zos.putNextEntry(ze);
        byte[] buffer = new byte[size];
        while ((length = inStream.read(buffer)) != -1) {
            zos.write(buffer, 0, length);
            zos.flush();
        }
        inStream.close();
        zos.closeEntry();
        initDocumentXmlRels();
        initContentTypes();
        this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk", fileId);
        if ("docx".equals(fileType)) {
            this.ContentTypes_xml.addInsertedDocx();
        } else if ("rtf".equals(fileType)) {
            this.ContentTypes_xml.addInsertedRtf();
        }
    }

    /** Helper methods to define document objects * */
    public Element getLineBreak() {
        Element r = getDocumentXmlDocument().createElement("w:r");
        Element cr = getDocumentXmlDocument().createElement("w:cr");
        r.appendChild(cr);
        return r;
    }

    public Element getWrapperElement() {
        return getDocumentXmlDocument().createElement("w:r");
    }

    public Element getBoldElement() {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element bold = getDocumentXmlDocument().createElement("w:b");
        rpr.appendChild(bold);
        return rpr;
    }

    public Element getItalicElement() {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element i = getDocumentXmlDocument().createElement("w:i");
        rpr.appendChild(i);
        return rpr;
    }

    public Element getTextElement(String text) {
        Element wrapper = getDocumentXmlDocument().createElement("w:rpr");
        Element el = getDocumentXmlDocument().createElement("w:t");
        el.setAttribute("xml:space", "preserve");
        Text t = getDocumentXmlDocument().createTextNode(text);
        el.appendChild(t);
        wrapper.appendChild(el);
        return wrapper;
    }

    public Element getParagraphElement() {
        return getParagraphElement(false);
    }

    public Element getParagraphElement(boolean singleSpace) {
        Element p = getDocumentXmlDocument().createElement("w:p");
        if (singleSpace) {
            Element ppr = this.document_xml.getDocument().createElement("w:pPr");
            Element w = this.document_xml.getDocument().createElement("w:spacing");
            w.setAttribute("w:line", "240");
            w.setAttribute("w:lineRule", "auto");
            ppr.appendChild(w);
            Element wcs = this.document_xml.getDocument().createElement("w:contextualSpacing");
            ppr.appendChild(wcs);
            p.appendChild(ppr);
        }
        return p;
    }

    public Element getColorElement(Color color) {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element clr = getDocumentXmlDocument().createElement("w:color");
        clr.setAttribute("w:val", convertColorToHexString(color));
        rpr.appendChild(clr);
        return rpr;
    }

    private String convertColorToHexString(Color color) {
        String retval = Integer.toHexString(color.getRGB() & 0x00ffffff);
        if (retval.length() == 5) {
            // TODO - find a better way to keep it from dropping leading 0
            retval = "0" + retval;
        }
        return retval;
    }

    public Element insertDocument(String name) {
        Element p = getParagraphElement(true);
        Element w = getWrapperElement();
        Element el = this.document_xml.getDocument().createElement("w:t");
        p.setAttribute("xml:space", "preserve");
        el.setAttribute("xml:space", "preserve");
        Element inserted = this.document_xml.getDocument().createElement("w:altChunk");
        inserted.setAttribute("r:id", name);
        el.appendChild(inserted);
        w.appendChild(el);
        p.appendChild(w);
        return p;
    }

    private Document getDocumentXml() throws Exception {
        if (this.document_xml == null)
            initDocumentXml();
        return this.document_xml.writeXml();
    }

    private Document getDocumentXmlDocument() {
        try {
            if (this.document_xml == null)
                initDocumentXml();
            return this.document_xml.getDocument();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private Document getFontTableXml() throws Exception {
        if (this.fontTable_xml == null)
            initFontTableXml();
        return this.fontTable_xml;
    }

    private Document getSettingsXml() throws Exception {
        if (this.settings_xml == null)
            initSettingsXml();
        return this.settings_xml;
    }

    private Document getStylesXml() throws Exception {
        if (this.styles_xml == null)
            initStylesXml();
        return this.styles_xml.writeXml();
    }

    private Document getWebSettingsXml() throws Exception {
        if (this.webSettings_xml == null)
            initWebSettingsXml();
        return this.webSettings_xml;
    }

    private Document getContentTypesXml() throws Exception {
        if (this.ContentTypes_xml == null)
            initContentTypes();
        return this.ContentTypes_xml.writeXml();
    }

    private Document getDocumentXmlRelsXml() throws Exception {
        if (this.document_xml_rels == null)
            initDocumentXmlRels();
        return this.document_xml_rels.writeXml();
    }

    private Document getThemeXml() throws Exception {
        if (this.theme_xml == null)
            initThemeXml();
        return this.theme_xml;
    }

    private Document getRelsXml() throws Exception {
        if (this.rels_xml == null)
            initRelsXml();
        return this.rels_xml.writeXml();
    }

    private Document getAppXml() throws Exception {
        if (this.app_xml == null)
            initAppXml();
        return this.app_xml;
    }

    private Document getCoreXml() throws Exception {
        if (this.core_xml == null)
            initCoreXml();
        return this.core_xml;
    }

    private void initDocumentXml() throws Exception {
        if (this.document_xml == null)
            this.document_xml = new DocumentXml();
    }

    // TODO - move into separate class
    private void initFontTableXml() throws Exception {
        this.fontTable_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.fontTable_xml.createElement("w:fonts");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        Element el;
        Element e;
        el = this.fontTable_xml.createElement("w:font");
        el.setAttribute("w:name", "Calibri");
        e = this.fontTable_xml.createElement("w:panose1");
        e.setAttribute("w:val", "020F0502020204030204");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:charset");
        e.setAttribute("w:val", "00");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:family");
        e.setAttribute("w:val", "swiss");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:pitch");
        e.setAttribute("w:val", "variable");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:sig");
        e.setAttribute("w:csb0", "0000009F");
        e.setAttribute("w:csb1", "00000000");
        e.setAttribute("w:usb0", "A00002EF");
        e.setAttribute("w:usb1", "4000207B");
        e.setAttribute("w:usb2", "00000000");
        e.setAttribute("w:usb3", "00000000");
        el.appendChild(e);
        root.appendChild(el);
        el = this.fontTable_xml.createElement("w:font");
        el.setAttribute("w:name", "Times New Roman");
        e = this.fontTable_xml.createElement("w:panose1");
        e.setAttribute("w:val", "02020603050405020304");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:charset");
        e.setAttribute("w:val", "00");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:family");
        e.setAttribute("w:val", "roman");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:pitch");
        e.setAttribute("w:val", "variable");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:sig");
        e.setAttribute("w:csb0", "000001FF");
        e.setAttribute("w:csb1", "00000000");
        e.setAttribute("w:usb0", "20002A87");
        e.setAttribute("w:usb1", "80000000");
        e.setAttribute("w:usb2", "00000008");
        e.setAttribute("w:usb3", "00000000");
        el.appendChild(e);
        root.appendChild(el);
        el = this.fontTable_xml.createElement("w:font");
        el.setAttribute("w:name", "Cambria");
        e = this.fontTable_xml.createElement("w:panose1");
        e.setAttribute("w:val", "02040503050406030204");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:charset");
        e.setAttribute("w:val", "00");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:family");
        e.setAttribute("w:val", "roman");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:pitch");
        e.setAttribute("w:val", "variable");
        el.appendChild(e);
        e = this.fontTable_xml.createElement("w:sig");
        e.setAttribute("w:csb0", "0000009F");
        e.setAttribute("w:csb1", "00000000");
        e.setAttribute("w:usb0", "A00002EF");
        e.setAttribute("w:usb1", "4000004B");
        e.setAttribute("w:usb2", "00000000");
        e.setAttribute("w:usb3", "00000000");
        el.appendChild(e);
        root.appendChild(el);
        this.fontTable_xml.appendChild(root);
    }

    // TODO - move into separate class
    private void initSettingsXml() throws Exception {
        this.settings_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.settings_xml.createElement("w:settings");
        root.setAttribute("xmlns:o", "urn:schemas-microsoft-com:office:office");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
        root.setAttribute("xmlns:v", "urn:schemas-microsoft-com:vml");
        root.setAttribute("xmlns:w10", "urn:schemas-microsoft-com:office:word");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        root.setAttribute("xmlns:sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
        Element el;
        Element e;
        el = this.settings_xml.createElement("w:zoom");
        el.setAttribute("w:percent", "100");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:defaultTabStop");
        el.setAttribute("w:val", "720");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:characterSpacingControl");
        el.setAttribute("w:val", "doNotCompress");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:compat");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:rsids");
        e = this.settings_xml.createElement("w:rsidRoot");
        e.setAttribute("w:val", "00993411");
        el.appendChild(e);
        e = this.settings_xml.createElement("w:rsid");
        e.setAttribute("w:val", "00993411");
        el.appendChild(e);
        e = this.settings_xml.createElement("w:rsid");
        e.setAttribute("w:val", "00C1102B");
        el.appendChild(e);
        e = this.settings_xml.createElement("w:rsid");
        e.setAttribute("w:val", "00C21D0A");
        el.appendChild(e);
        e = this.settings_xml.createElement("w:rsid");
        e.setAttribute("w:val", "00E827D4");
        el.appendChild(e);
        root.appendChild(el);
        el = this.settings_xml.createElement("m:mathPr");
        e = this.settings_xml.createElement("m:mathFont");
        e.setAttribute("m:val", "Cambria Math");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:brkBin");
        e.setAttribute("m:val", "before");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:brkBinSub");
        e.setAttribute("m:val", "--");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:smallFrac");
        e.setAttribute("m:val", "off");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:dispDef");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:lMargin");
        e.setAttribute("m:val", "0");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:rMargin");
        e.setAttribute("m:val", "0");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:defJc");
        e.setAttribute("m:val", "centerGroup");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:wrapIndent");
        e.setAttribute("m:val", "1440");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:intLim");
        e.setAttribute("m:val", "subSup");
        el.appendChild(e);
        e = this.settings_xml.createElement("m:naryLim");
        e.setAttribute("m:val", "undOvr");
        el.appendChild(e);
        root.appendChild(el);
        el = this.settings_xml.createElement("w:themeFontLang");
        el.setAttribute("w:val", "en-US");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:clrSchemeMapping");
        el.setAttribute("w:accent1", "accent1");
        el.setAttribute("w:accent2", "accent2");
        el.setAttribute("w:accent3", "accent3");
        el.setAttribute("w:accent4", "accent4");
        el.setAttribute("w:accent5", "accent5");
        el.setAttribute("w:accent6", "accent6");
        el.setAttribute("w:bg1", "light1");
        el.setAttribute("w:bg2", "light2");
        el.setAttribute("w:followedHyperlink", "followedHyperlink");
        el.setAttribute("w:hyperlink", "hyperlink");
        el.setAttribute("w:t1", "dark1");
        el.setAttribute("w:t2", "dark2");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:shapeDefaults");
        e = this.settings_xml.createElement("o:shapedefaults");
        e.setAttribute("spidmax", "2050");
        e.setAttribute("v:ext", "edit");
        el.appendChild(e);
        e = this.settings_xml.createElement("o:shapelayout");
        e.setAttribute("v:ext", "edit");
        e = this.settings_xml.createElement("o:idmap");
        e.setAttribute("data", "1");
        e.setAttribute("v:ext", "edit");
        el.appendChild(e);
        el.appendChild(e);
        root.appendChild(el);
        el = this.settings_xml.createElement("w:decimalSymbol");
        el.setAttribute("w:val", ".");
        root.appendChild(el);
        el = this.settings_xml.createElement("w:listSeparator");
        el.setAttribute("w:val", ",");
        root.appendChild(el);
        this.settings_xml.appendChild(root);
    }

    // TODO - move into separate class
    private void initStylesXml() throws Exception {
        initContentTypes();
        initDocumentXmlRels();
        this.ContentTypes_xml.addOverride("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
        this.document_xml_rels.addRelationship("styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        if (this.styles_xml == null)
            this.styles_xml = new Styles();
    }

    // TODO - move into separate class
    private void initWebSettingsXml() throws Exception {
        this.webSettings_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.webSettings_xml.createElement("w:webSettings");
        root.setAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        root.setAttribute("xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        Element el;
        el = this.webSettings_xml.createElement("w:optimizeForBrowser");
        root.appendChild(el);
        this.webSettings_xml.appendChild(root);
    }

    private void initContentTypes() throws Exception {
        if (this.ContentTypes_xml == null)
            this.ContentTypes_xml = new ContentTypes();
    }

    private void initDocumentXmlRels() throws Exception {
        if (this.document_xml_rels == null)
            this.document_xml_rels = new DocumentXmlRels();
    }

    // TODO - move into separate class
    private void initThemeXml() throws Exception {
        this.theme_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.theme_xml.createElement("a:theme");
        root.setAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        root.setAttribute("name", "Office Theme");
        Element el;
        Element e;
        el = this.theme_xml.createElement("a:themeElements");
        root.appendChild(el);
        e = this.theme_xml.createElement("a:clrScheme");
        e.setAttribute("name", "Office");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:dk1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:sysClr");
        e.setAttribute("lastClr", "000000");
        e.setAttribute("val", "windowText");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lt1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:sysClr");
        e.setAttribute("lastClr", "FFFFFF");
        e.setAttribute("val", "window");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:dk2");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "1F497D");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lt2");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "EEECE1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "4F81BD");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent2");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "C0504D");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent3");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "9BBB59");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent4");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "8064A2");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent5");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "4BACC6");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:accent6");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "F79646");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:hlink");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "0000FF");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:folHlink");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "800080");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:fontScheme");
        e.setAttribute("name", "Office");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:majorFont");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:latin");
        e.setAttribute("typeface", "Cambria");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:ea");
        e.setAttribute("typeface", "");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:cs");
        e.setAttribute("typeface", "");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Jpan");
        e.setAttribute("typeface", "?? ????");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hang");
        e.setAttribute("typeface", "?? ??");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hans");
        e.setAttribute("typeface", "??");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hant");
        e.setAttribute("typeface", "????");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Arab");
        e.setAttribute("typeface", "Times New Roman");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hebr");
        e.setAttribute("typeface", "Times New Roman");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Thai");
        e.setAttribute("typeface", "Angsana New");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Ethi");
        e.setAttribute("typeface", "Nyala");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Beng");
        e.setAttribute("typeface", "Vrinda");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Gujr");
        e.setAttribute("typeface", "Shruti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Khmr");
        e.setAttribute("typeface", "MoolBoran");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Knda");
        e.setAttribute("typeface", "Tunga");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Guru");
        e.setAttribute("typeface", "Raavi");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Cans");
        e.setAttribute("typeface", "Euphemia");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Cher");
        e.setAttribute("typeface", "Plantagenet Cherokee");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Yiii");
        e.setAttribute("typeface", "Microsoft Yi Baiti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Tibt");
        e.setAttribute("typeface", "Microsoft Himalaya");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Thaa");
        e.setAttribute("typeface", "MV Boli");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Deva");
        e.setAttribute("typeface", "Mangal");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Telu");
        e.setAttribute("typeface", "Gautami");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Taml");
        e.setAttribute("typeface", "Latha");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Syrc");
        e.setAttribute("typeface", "Estrangelo Edessa");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Orya");
        e.setAttribute("typeface", "Kalinga");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Mlym");
        e.setAttribute("typeface", "Kartika");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Laoo");
        e.setAttribute("typeface", "DokChampa");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Sinh");
        e.setAttribute("typeface", "Iskoola Pota");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Mong");
        e.setAttribute("typeface", "Mongolian Baiti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Viet");
        e.setAttribute("typeface", "Times New Roman");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Uigh");
        e.setAttribute("typeface", "Microsoft Uighur");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:minorFont");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:latin");
        e.setAttribute("typeface", "Calibri");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:ea");
        e.setAttribute("typeface", "");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:cs");
        e.setAttribute("typeface", "");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Jpan");
        e.setAttribute("typeface", "?? ??");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hang");
        e.setAttribute("typeface", "?? ??");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hans");
        e.setAttribute("typeface", "??");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hant");
        e.setAttribute("typeface", "????");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Arab");
        e.setAttribute("typeface", "Arial");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Hebr");
        e.setAttribute("typeface", "Arial");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Thai");
        e.setAttribute("typeface", "Cordia New");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Ethi");
        e.setAttribute("typeface", "Nyala");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Beng");
        e.setAttribute("typeface", "Vrinda");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Gujr");
        e.setAttribute("typeface", "Shruti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Khmr");
        e.setAttribute("typeface", "DaunPenh");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Knda");
        e.setAttribute("typeface", "Tunga");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Guru");
        e.setAttribute("typeface", "Raavi");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Cans");
        e.setAttribute("typeface", "Euphemia");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Cher");
        e.setAttribute("typeface", "Plantagenet Cherokee");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Yiii");
        e.setAttribute("typeface", "Microsoft Yi Baiti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Tibt");
        e.setAttribute("typeface", "Microsoft Himalaya");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Thaa");
        e.setAttribute("typeface", "MV Boli");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Deva");
        e.setAttribute("typeface", "Mangal");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Telu");
        e.setAttribute("typeface", "Gautami");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Taml");
        e.setAttribute("typeface", "Latha");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Syrc");
        e.setAttribute("typeface", "Estrangelo Edessa");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Orya");
        e.setAttribute("typeface", "Kalinga");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Mlym");
        e.setAttribute("typeface", "Kartika");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Laoo");
        e.setAttribute("typeface", "DokChampa");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Sinh");
        e.setAttribute("typeface", "Iskoola Pota");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Mong");
        e.setAttribute("typeface", "Mongolian Baiti");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Viet");
        e.setAttribute("typeface", "Arial");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:font");
        e.setAttribute("script", "Uigh");
        e.setAttribute("typeface", "Microsoft Uighur");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:fmtScheme");
        e.setAttribute("name", "Office");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:fillStyleLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:solidFill");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gradFill");
        e.setAttribute("rotWithShape", "1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gsLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "50000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "300000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "35000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "37000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "300000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "100000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "15000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "350000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lin");
        e.setAttribute("ang", "16200000");
        e.setAttribute("scaled", "1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gradFill");
        e.setAttribute("rotWithShape", "1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gsLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "51000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "130000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "80000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "93000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "130000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "100000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "94000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "135000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lin");
        e.setAttribute("ang", "16200000");
        e.setAttribute("scaled", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lnStyleLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:ln");
        e.setAttribute("algn", "ctr");
        e.setAttribute("cap", "flat");
        e.setAttribute("cmpd", "sng");
        e.setAttribute("w", "9525");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:solidFill");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "95000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "105000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:prstDash");
        e.setAttribute("val", "solid");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:ln");
        e.setAttribute("algn", "ctr");
        e.setAttribute("cap", "flat");
        e.setAttribute("cmpd", "sng");
        e.setAttribute("w", "25400");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:solidFill");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:prstDash");
        e.setAttribute("val", "solid");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:ln");
        e.setAttribute("algn", "ctr");
        e.setAttribute("cap", "flat");
        e.setAttribute("cmpd", "sng");
        e.setAttribute("w", "38100");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:solidFill");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:prstDash");
        e.setAttribute("val", "solid");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectStyleLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectStyle");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:outerShdw");
        e.setAttribute("blurRad", "40000");
        e.setAttribute("dir", "5400000");
        e.setAttribute("dist", "20000");
        e.setAttribute("rotWithShape", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "000000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:alpha");
        e.setAttribute("val", "38000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectStyle");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:outerShdw");
        e.setAttribute("blurRad", "40000");
        e.setAttribute("dir", "5400000");
        e.setAttribute("dist", "23000");
        e.setAttribute("rotWithShape", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "000000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:alpha");
        e.setAttribute("val", "35000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectStyle");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:effectLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:outerShdw");
        e.setAttribute("blurRad", "40000");
        e.setAttribute("dir", "5400000");
        e.setAttribute("dist", "23000");
        e.setAttribute("rotWithShape", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:srgbClr");
        e.setAttribute("val", "000000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:alpha");
        e.setAttribute("val", "35000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:scene3d");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:camera");
        e.setAttribute("prst", "orthographicFront");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:rot");
        e.setAttribute("lat", "0");
        e.setAttribute("lon", "0");
        e.setAttribute("rev", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:lightRig");
        e.setAttribute("dir", "t");
        e.setAttribute("rig", "threePt");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:rot");
        e.setAttribute("lat", "0");
        e.setAttribute("lon", "0");
        e.setAttribute("rev", "1200000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:sp3d");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:bevelT");
        e.setAttribute("h", "25400");
        e.setAttribute("w", "63500");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:bgFillStyleLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:solidFill");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gradFill");
        e.setAttribute("rotWithShape", "1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gsLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "40000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "350000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "40000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "45000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "99000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "350000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "100000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "20000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "255000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:path");
        e.setAttribute("path", "circle");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:fillToRect");
        e.setAttribute("b", "180000");
        e.setAttribute("l", "50000");
        e.setAttribute("r", "50000");
        e.setAttribute("t", "-80000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gradFill");
        e.setAttribute("rotWithShape", "1");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gsLst");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "0");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:tint");
        e.setAttribute("val", "80000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "300000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:gs");
        e.setAttribute("pos", "100000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:schemeClr");
        e.setAttribute("val", "phClr");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:shade");
        e.setAttribute("val", "30000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:satMod");
        e.setAttribute("val", "200000");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:path");
        e.setAttribute("path", "circle");
        el.appendChild(e);
        e = this.theme_xml.createElement("a:fillToRect");
        e.setAttribute("b", "50000");
        e.setAttribute("l", "50000");
        e.setAttribute("r", "50000");
        e.setAttribute("t", "50000");
        el.appendChild(e);
        el = this.theme_xml.createElement("a:objectDefaults");
        root.appendChild(el);
        el = this.theme_xml.createElement("a:extraClrSchemeLst");
        root.appendChild(el);
        this.theme_xml.appendChild(root);
    }

    private void initRelsXml() throws Exception {
        if (this.rels_xml == null)
            this.rels_xml = new RelsRels();
    }

    // TODO - move into separate class
    private void initAppXml() throws Exception {
        this.app_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.app_xml.createElement("Properties");
        root.setAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        root.setAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        Element el;
        Text t;
        el = this.app_xml.createElement("Template");
        t = this.app_xml.createTextNode("Normal.dotm");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("TotalTime");
        t = this.app_xml.createTextNode("0");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Pages");
        t = this.app_xml.createTextNode("1");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Words");
        t = this.app_xml.createTextNode("0");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Characters");
        t = this.app_xml.createTextNode("0");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Application");
        t = this.app_xml.createTextNode("Microsoft Office Word");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("DocSecurity");
        t = this.app_xml.createTextNode("0");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Lines");
        t = this.app_xml.createTextNode("1");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Paragraphs");
        t = this.app_xml.createTextNode("1");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("ScaleCrop");
        t = this.app_xml.createTextNode("false");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("Company");
        root.appendChild(el);
        el = this.app_xml.createElement("LinksUpToDate");
        t = this.app_xml.createTextNode("false");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("CharactersWithSpaces");
        t = this.app_xml.createTextNode("0");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("SharedDoc");
        t = this.app_xml.createTextNode("false");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("HyperlinksChanged");
        t = this.app_xml.createTextNode("false");
        el.appendChild(t);
        root.appendChild(el);
        el = this.app_xml.createElement("AppVersion");
        t = this.app_xml.createTextNode("12.0000");
        el.appendChild(t);
        root.appendChild(el);
        this.app_xml.appendChild(root);
    }

    // TODO - move into separate class
    private void initCoreXml() throws Exception {
        this.core_xml = dbf.newDocumentBuilder().newDocument();
        Element root = this.core_xml.createElement("cp:coreProperties");
        root.setAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        root.setAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/");
        root.setAttribute("xmlns:dcterms", "http://purl.org/dc/terms/");
        root.setAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
        root.setAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
        Element el;
        Text t;
        el = this.core_xml.createElement("dc:title");
        t = this.core_xml.createTextNode("Test");
        el.appendChild(t);
        root.appendChild(el);
        el = this.core_xml.createElement("dc:subject");
        root.appendChild(el);
        el = this.core_xml.createElement("dc:creator");
        t = this.core_xml.createTextNode("Budget Formulation and Exection Manager");
        el.appendChild(t);
        root.appendChild(el);
        el = this.core_xml.createElement("cp:keywords");
        root.appendChild(el);
        el = this.core_xml.createElement("dc:description");
        root.appendChild(el);
        el = this.core_xml.createElement("cp:lastModifiedBy");
        t = this.core_xml.createTextNode("Budget Formulation and Exection Manage");
        el.appendChild(t);
        root.appendChild(el);
        el = this.core_xml.createElement("cp:revision");
        t = this.core_xml.createTextNode("1");
        el.appendChild(t);
        root.appendChild(el);
        el = this.core_xml.createElement("dcterms:created");
        el.setAttribute("xsi:type", "dcterms:W3CDTF");
        java.util.Date now = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-ddHH:mm:ss");
        t = this.core_xml.createTextNode(sdf.format(now));
        el.appendChild(t);
        root.appendChild(el);
        el = this.core_xml.createElement("dcterms:modified");
        el.setAttribute("xsi:type", "dcterms:W3CDTF");
        t = this.core_xml.createTextNode(sdf.format(now));
        el.appendChild(t);
        root.appendChild(el);
        this.core_xml.appendChild(root);
    }

    private void saveDebug() throws Exception {
        addEntry(zos, getDocumentXml(), "word/document.xml");
        addEntry(zos, getContentTypesXml(), "[Content_Types].xml");
        addEntry(zos, getRelsXml(), "_rels/.rels");
        if (this.document_xml_rels != null) {
            addEntry(zos, getDocumentXmlRelsXml(), "word/_rels/document.xml.rels");
        }
        addEntry(zos, getStylesXml(), "word/styles.xml");
        if (this.ADD_OPTIONAL_XML) {
            // NB - this doesn't quite work yet
            addEntry(zos, getFontTableXml(), "word/fontTable.xml");
            addEntry(zos, getSettingsXml(), "word/settings.xml");
            addEntry(zos, getWebSettingsXml(), "word/webSettings.xml");
            addEntry(zos, getThemeXml(), "word/theme/theme1.xml");
            addEntry(zos, getAppXml(), "docProps/app.xml");
            addEntry(zos, getCoreXml(), "docProps/core.xml");
        }
        zos.flush();
        zos.close();
    }

    public BLOB save() throws Exception {
        if (this.isDebug) {
            saveDebug();
            return null;
        }
        Connection conn = Utils.openConnection();
        addEntry(zos, getDocumentXml(), "word/document.xml");
        addEntry(zos, getContentTypesXml(), "[Content_Types].xml");
        addEntry(zos, getRelsXml(), "_rels/.rels");
        if (this.document_xml_rels != null) {
            addEntry(zos, getDocumentXmlRelsXml(), "word/_rels/document.xml.rels");
        }
        if (this.styles_xml != null) {
            addEntry(zos, getStylesXml(), "word/styles.xml");
        }
        if (this.ADD_OPTIONAL_XML) {
            // NB - this doesn't quite work yet
            addEntry(zos, getFontTableXml(), "word/fontTable.xml");
            addEntry(zos, getSettingsXml(), "word/settings.xml");
            addEntry(zos, getWebSettingsXml(), "word/webSettings.xml");
            addEntry(zos, getThemeXml(), "word/theme/theme1.xml");
            addEntry(zos, getAppXml(), "docProps/app.xml");
            addEntry(zos, getCoreXml(), "docProps/core.xml");
        }
        zos.flush();
        zos.close();
        BLOB retval = BLOB.createTemporary(conn, false, BLOB.DURATION_SESSION);
        Utils.zipBlob("this.docx", this.tempBlob, retval);
        BLOB.freeTemporary(this.tempBlob);
        return retval;
    }

    private void addEntry(ZipOutputStream zos, Document doc, String fileName) throws Exception {
        Transformer t = TransformerFactory.newInstance().newTransformer();
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        t.transform(new DOMSource(doc), new StreamResult(baos));
        ZipEntry ze = new ZipEntry(fileName);
        byte[] data = baos.toByteArray();
        ze.setSize(data.length);
        zos.putNextEntry(ze);
        zos.write(data);
        zos.flush();
        zos.closeEntry();
    }

    public static void main(String[] args) {
        try {
            Docx doc = new Docx("c:\\temp\\mytest.docx");
            Element p = doc.getParagraphElement();
            Element w = doc.getWrapperElement();
            Element t = doc.getTextElement("This is a test document.");
            w.appendChild(t);
            p.appendChild(w);
            doc.appendChild(p);
            doc.save();
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
