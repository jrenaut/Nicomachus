package docx;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import oracle.sql.BLOB;
import oracle.sql.CLOB;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

public class Docx {
    private DocumentXml     document_xml;
    private FontTable       fontTable_xml;
    private Settings        settings_xml;
    private Styles          styles_xml;
    private WebSettings     webSettings_xml;
    private ContentTypes    ContentTypes_xml;
    private DocumentXmlRels document_xml_rels;
    private Theme           theme_xml;
    private RelsRels        rels_xml;
    private App             app_xml;
    private Core            core_xml;
    // NB - if you set ADD_OPTIONAL_XML to true, you get an invalid docx file.
    // Still working
    // on that. -- JER
    private boolean         ADD_OPTIONAL_XML = false;
    private ZipOutputStream zos;
    private BLOB            tempBlob;
    private boolean         isDebug          = false;
    private int             tocElementCount  = 1;

    public Docx() {
        try {
            Connection conn = Utils.openConnection();
            this.tempBlob = BLOB.createTemporary(conn, false, BLOB.DURATION_SESSION);
            this.tempBlob.open(BLOB.MODE_READWRITE);
            this.tempBlob.truncate(0);
            this.zos = new ZipOutputStream(tempBlob.setBinaryStream(1L));
        } catch (Exception e) {
            //
        }
    }

    public Docx(String debugFileName) {
        this.isDebug = true;
        try {
            this.zos = new ZipOutputStream(new FileOutputStream(debugFileName));
        } catch (Exception e) {
            //
        }
    }

    public void appendChild(Element e) {
        getDocumentBody().appendChild(e);
    }

    private Element getDocumentBody() {
        try {
            initDocumentXml();
            return this.document_xml.getBody();
        } catch (Exception e) {
            return null;
        }
    }

    public void addInsertedDocument(BLOB b, String fileType, String fileId) throws Exception {
        String filename = generateFilename(fileType);
        Connection conn = Utils.openConnection();
        BLOB temp = Utils.createTemporaryBlob(conn);
        Utils.unzipBlobtoBlob("", b, temp);
        if ("doc".equals(fileType)) {
            filename = "embeddings/" + filename;
        } else {
            filename = "word/" + filename;
        }
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
        if ("docx".equals(fileType)) {
            this.ContentTypes_xml.addInsertedDocx();
        } else if ("rtf".equals(fileType)) {
            this.ContentTypes_xml.addInsertedRtf();
        } else if ("xls".equals(fileType)) {
            this.ContentTypes_xml.addInsertedXls();
        } else if ("doc".equals(fileType)) {
            throw new Exception("Unsupported filetype: doc");
            // this.ContentTypes_xml.addInsertedDoc();
        } else {
            throw new Exception("Unsupported filetype: " + fileType);
        }
        if ("doc".equals(fileType)) {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", fileId);
        } else {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk", fileId);
        }
    }

    public void addInsertedDocument(byte[] b, String fileType, String fileId) throws Exception {
        String filename = generateFilename(fileType);
        Connection conn = Utils.openConnection();
        BLOB temp = Utils.createTemporaryBlob(conn);
        unzipBlobtoBlob("", b, temp);
        if ("doc".equals(fileType)) {
            filename = "embeddings/" + filename;
        } else {
            filename = "word/" + filename;
        }
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
        if ("docx".equals(fileType)) {
            this.ContentTypes_xml.addInsertedDocx();
        } else if ("rtf".equals(fileType)) {
            this.ContentTypes_xml.addInsertedRtf();
        } else if ("xls".equals(fileType)) {
            this.ContentTypes_xml.addInsertedXls();
        } else if ("doc".equals(fileType)) {
            throw new Exception("Unsupported filetype: doc");
            // this.ContentTypes_xml.addInsertedDoc();
        } else {
            throw new Exception("Unsupported filetype: " + fileType);
        }
        if ("doc".equals(fileType)) {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", fileId);
        } else {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk", fileId);
        }
    }

    /**
     * Ensure unique filename for inserted documents
     * 
     * @param ext
     * @return
     */
    private String generateFilename(String ext) {
        Date d = new Date();
        return "file" + d.getTime() + "." + ext;
    }

    private void unzipBlobtoBlob(String fileName, byte[] input, BLOB output) throws SQLException, IOException {
        InputStream is = new ByteArrayInputStream(input);
        OutputStream out = output.setBinaryStream(1L);
        ZipInputStream in = new ZipInputStream(is);
        in.getNextEntry();
        byte[] buf = new byte[1024];
        byte[] buf2;
        int len;
        while ((len = in.read(buf)) > 0) {
            buf2 = new byte[len];
            System.arraycopy(buf, 0, buf2, 0, len);
            out.write(buf2);
        }
        out.close();
        in.close();
    }

    public void addInsertedUnzippedDocument(BLOB b, String fileType, String fileId) throws Exception {
        String filename = generateFilename(fileType);
        if ("doc".equals(fileType)) {
            filename = "embeddings/" + filename;
        } else {
            filename = "word/" + filename;
        }
        ZipEntry ze = new ZipEntry(filename);
        InputStream inStream = b.binaryStreamValue();
        int length = -1;
        int size = b.getBufferSize();
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
        if ("docx".equals(fileType)) {
            this.ContentTypes_xml.addInsertedDocx();
        } else if ("rtf".equals(fileType)) {
            this.ContentTypes_xml.addInsertedRtf();
        } else if ("xls".equals(fileType)) {
            this.ContentTypes_xml.addInsertedXls();
        } else if ("doc".equals(fileType)) {
            throw new Exception("Unsupported filetype: doc");
            // this.ContentTypes_xml.addInsertedDoc();
        } else {
            throw new Exception("Unsupported filetype: " + fileType);
        }
        if ("doc".equals(fileType)) {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", fileId);
        } else {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk", fileId);
        }
    }

    public void addInsertedUnzippedDocument(CLOB c, String fileType, String fileId) throws Exception {
        String filename = generateFilename(fileType);
        filename = "word/" + filename;
        ZipEntry ze = new ZipEntry(filename);
        InputStream inStream = c.binaryStreamValue();
        int length = -1;
        int size = c.getBufferSize();
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
        if ("docx".equals(fileType)) {
            this.ContentTypes_xml.addInsertedDocx();
        } else if ("rtf".equals(fileType)) {
            this.ContentTypes_xml.addInsertedRtf();
        } else if ("xls".equals(fileType)) {
            this.ContentTypes_xml.addInsertedXls();
        } else if ("doc".equals(fileType)) {
            throw new Exception("Unsupported filetype: doc");
            // this.ContentTypes_xml.addInsertedDoc();
        } else {
            throw new Exception("Unsupported filetype: " + fileType);
        }
        if ("doc".equals(fileType)) {
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", fileId);
        } else {
            // this.document_xml_rels.addSubDocument("/" + filename,
            // "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument",
            // fileId);
            this.document_xml_rels.addRelationship("/" + filename, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk", fileId);
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

    public Element getUnderlineElement() {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element u = getDocumentXmlDocument().createElement("w:u");
        u.setAttribute("w:val", "single");
        rpr.appendChild(u);
        return rpr;
    }

    public Element getItalicElement() {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element i = getDocumentXmlDocument().createElement("w:i");
        rpr.appendChild(i);
        return rpr;
    }

    public Element getTextElement(String text) {
        Element wrapper = getDocumentXmlDocument().createElement("w:rPr");
        Element el = getDocumentXmlDocument().createElement("w:t");
        el.setAttribute("xml:space", "preserve");
        Text t = getDocumentXmlDocument().createTextNode(text);
        el.appendChild(t);
        wrapper.appendChild(el);
        return wrapper;
    }

    public Element getTextElement(String text, boolean noPr) {
        Element wrapper = getDocumentXmlDocument().createElement("w:r");
        Element el = getDocumentXmlDocument().createElement("w:t");
        el.setAttribute("xml:space", "preserve");
        Text t = getDocumentXmlDocument().createTextNode(text);
        el.appendChild(t);
        wrapper.appendChild(el);
        return wrapper;
    }

    /**
     * TODO - needs to be abstract and generic
     * 
     * @param text
     * @return
     * @throws Exception
     */
    public Element getHeaderElement(String text) throws Exception {
        if (this.tocElementCount == 0) {
            getStyles().addDefaultStyle();
        }
        Element p = getParagraphElement();
        Element pPr = getDocumentXmlDocument().createElement("w:pPr");
        Element pStyle = getDocumentXmlDocument().createElement("w:pStyle");
        pStyle.setAttribute("w:val", "HeadingCustom");
        pPr.appendChild(pStyle);
        p.appendChild(pPr);
        Element bk = getDocumentXmlDocument().createElement("w:bookmarkStart");
        bk.setAttribute("w:id", "" + (this.tocElementCount - 1));
        bk.setAttribute("w:name", "_TocHeading" + this.tocElementCount);
        p.appendChild(bk);
        Element t = getTextElement(text, true);
        p.appendChild(t);
        bk = getDocumentXmlDocument().createElement("w:bookmarkEnd");
        bk.setAttribute("w:id", "" + (this.tocElementCount - 1));
        p.appendChild(bk);
        this.tocElementCount++;
        return p;
    }

    public Element getParagraphElement() {
        return getParagraphElement(false);
    }

    public Element getParagraphElement(boolean singleSpace) {
        Element p = getDocumentXmlDocument().createElement("w:p");
        if (singleSpace) {
            Element ppr = getDocumentXmlDocument().createElement("w:pPr");
            Element w = getDocumentXmlDocument().createElement("w:spacing");
            w.setAttribute("w:line", "240");
            w.setAttribute("w:lineRule", "auto");
            ppr.appendChild(w);
            Element wcs = getDocumentXmlDocument().createElement("w:contextualSpacing");
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

    public Element getAlignmentElement(String alignment) {
        Element ppr = getDocumentXmlDocument().createElement("w:pPr");
        Element align = getDocumentXmlDocument().createElement("w:jc");
        align.setAttribute("w:val", alignment);
        ppr.appendChild(align);
        return ppr;
    }

    public Element getFontElement(String fontname, int fontsize) {
        Element rpr = getDocumentXmlDocument().createElement("w:rPr");
        Element rFonts = getDocumentXmlDocument().createElement("w:rFonts");
        rFonts.setAttribute("w:ascii", fontname);
        rFonts.setAttribute("w:hAnsi", fontname);
        rFonts.setAttribute("w:cs", fontname);
        rpr.appendChild(rFonts);
        Element sz = getDocumentXmlDocument().createElement("w:sz");
        sz.setAttribute("w:val", "" + (2 * fontsize));
        rpr.appendChild(sz);
        Element szCs = getDocumentXmlDocument().createElement("w:szCs");
        szCs.setAttribute("w:val", "" + (2 * fontsize));
        rpr.appendChild(szCs);
        return rpr;
    }

    public Element insertDocument(String name) {
        Element p = getParagraphElement(true);
        Element w = getWrapperElement();
        Element el = getDocumentXmlDocument().createElement("w:t");
        p.setAttribute("xml:space", "preserve");
        el.setAttribute("xml:space", "preserve");
        Element inserted = getDocumentXmlDocument().createElement("w:altChunk");
        inserted.setAttribute("r:id", name);
        el.appendChild(inserted);
        w.appendChild(el);
        p.appendChild(w);
        return p;
    }

    public Element insertTable(Table table) throws Exception {
        Element t = getDocumentXmlDocument().createElement("w:tbl");
        Element tblPr = getDocumentXmlDocument().createElement("w:tblPr");
        if (table.keepOnOnePage()) {
            getStyles().addTableGrid();
            Element tblStyle = getDocumentXmlDocument().createElement("w:tblStyle");
            tblStyle.setAttribute("w:val", "TableOnePage");
            tblPr.appendChild(tblStyle);
        }
        Element tblw = getDocumentXmlDocument().createElement("w:tblW");
        tblw.setAttribute("w:w", "" + table.getWidth());
        if (table.getWidth() > 0) {
            tblw.setAttribute("w:type", "dxa");
        } else {
            tblw.setAttribute("w:type", "auto");
        }
        tblPr.appendChild(tblw);
        t.appendChild(tblPr);
        Element tgrid = getDocumentXmlDocument().createElement("w:tblGrid");
        double[] cols = table.getColumnWidths();
        if (cols != null) {
            for (int i = 0; i < cols.length; i++) {
                Element gridCol = getDocumentXmlDocument().createElement("w:gridCol");
                gridCol.setAttribute("w:w", "" + cols[i]);
                tgrid.appendChild(gridCol);
            }
        }
        t.appendChild(tgrid);
        for (Iterator it = table.getRows().iterator(); it.hasNext();) {
            Row r = (Row) it.next();
            if (table.getRowHeight() != 0 && r.getRowHeight() == 0) {
                r.setRowHeight(table.getRowHeight());
            }
            t.appendChild(addTableRow(r));
        }
        return t;
    }

    private Element addTableRow(Row r) {
        Element row = getDocumentXmlDocument().createElement("w:tr");
        if (r.needsTrPr()) {
            Element trPr = getDocumentXmlDocument().createElement("w:trPr");
            if (!r.canBreakAcrossPages()) {
                Element cs = getDocumentXmlDocument().createElement("w:cantSplit");
                trPr.appendChild(cs);
            }
            if (r.getRowHeight() != 0) {
                Element h = getDocumentXmlDocument().createElement("w:trHeight");
                h.setAttribute("w:hRule", "exact");
                h.setAttribute("w:val", "" + (r.getRowHeight()));
                trPr.appendChild(h);
            }
            row.appendChild(trPr);
        }
        for (Iterator it = r.getCells().iterator(); it.hasNext();) {
            Cell c = (Cell) it.next();
            row.appendChild(addTableCell(c));
        }
        return row;
    }

    private Element addTableCell(Cell c) {
        Element cell = getDocumentXmlDocument().createElement("w:tc");
        Element tcpr = getDocumentXmlDocument().createElement("w:tcPr");
        Element shd = getDocumentXmlDocument().createElement("w:shd");
        shd.setAttribute("w:val", "clear");
        shd.setAttribute("w:color", "auto");
        shd.setAttribute("w:fill", convertColorToHexString(c.getBackgroundColor()));
        tcpr.appendChild(shd);
        if (c.getWidth() != -1) {
            Element tcw = getDocumentXmlDocument().createElement("w:tcW");
            tcw.setAttribute("w:w", "" + c.getWidth());
            tcw.setAttribute("w:type", "dxa");
            tcpr.appendChild(tcw);
        }
        if (c.getRowSpan() > 1) {
            // TODO Much more complicated, not going to implement now
        }
        if (c.getColumnSpan() > 1) {
            Element wgridspan = getDocumentXmlDocument().createElement("w:gridSpan");
            wgridspan.setAttribute("w:val", "" + c.getColumnSpan());
            tcpr.appendChild(wgridspan);
        }
        ArrayList borders = c.getBorders();
        if (borders != null) {
            Element pBdr = getDocumentXmlDocument().createElement("w:pBdr");
            for (Iterator it = borders.iterator(); it.hasNext();) {
                CellBorder cb = (CellBorder) it.next();
                Element bdr = getDocumentXmlDocument().createElement("w:" + cb.getLocation());
                bdr.setAttribute("w:val", cb.getType());
                bdr.setAttribute("w:sz", "" + cb.getSize());
                bdr.setAttribute("w:color", "auto");
                pBdr.appendChild(bdr);
            }
            tcpr.appendChild(pBdr);
        }
        cell.appendChild(tcpr);
        Element p = getParagraphElement();
        Element w = getWrapperElement();
        w.appendChild(getColorElement(c.getFontColor()));
        p.appendChild(getAlignmentElement(c.getAlignment()));
        w.appendChild(getFontElement(c.getFontName(), c.getFontSize()));
        if (c.isBold()) {
            w.appendChild(getBoldElement());
        }
        if (c.isItalic()) {
            w.appendChild(getItalicElement());
        }
        Element t = getTextElement(c.getText());
        w.appendChild(t);
        p.appendChild(w);
        cell.appendChild(p);
        return cell;
    }

    /**
     * DOES NOT WORK
     * 
     * @param name
     * @return
     */
    public Element insertSubDocument(String name) {
        Element subdoc = getDocumentXmlDocument().createElement("w:subDoc");
        subdoc.setAttribute("r:id", name);
        return subdoc;
    }

    /**
     * DOES NOT WORK
     * 
     * @param name
     * @return
     */
    public Element insertDocDocument(String name) {
        // ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1331581368"
        // r:id="rId6">
        Element p = getParagraphElement(true);
        Element w = getWrapperElement();
        Element o = getDocumentXmlDocument().createElement("w:object");
        Element ole = getDocumentXmlDocument().createElement("o:OLEObject");
        ole.setAttribute("Type", "Embed");
        ole.setAttribute("ProgID", "Word.Document.8");
        ole.setAttribute("DrawAspect", "Content");
        ole.setAttribute("ObjectID", name);
        o.appendChild(ole);
        w.appendChild(o);
        p.appendChild(w);
        return p;
    }

    public Element getDotLeaders() {
        double defaultSize = 6.5;
        return getDotLeaders("right", "dot", "" + (1440 * defaultSize));
    }

    public Element getTab() {
        return getDocumentXmlDocument().createElement("w:tab");
    }

    public Element getDotLeaders(String alignment, String leader, String position) {
        Element pPr = getDocumentXmlDocument().createElement("w:pPr");
        Element tabs = getDocumentXmlDocument().createElement("w:tabs");
        Element tab = getDocumentXmlDocument().createElement("w:tab");
        tab.setAttribute("w:val", alignment);
        if (leader != null) {
            tab.setAttribute("w:leader", leader);
        }
        tab.setAttribute("w:pos", position);
        tabs.appendChild(tab);
        pPr.appendChild(tabs);
        return pPr;
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
            return null;
        }
    }

    private Document getFontTableXml() throws Exception {
        if (this.fontTable_xml == null)
            initFontTableXml();
        return this.fontTable_xml.writeXml();
    }

    private Document getSettingsXml() throws Exception {
        if (this.settings_xml == null)
            initSettingsXml();
        return this.settings_xml.writeXml();
    }

    private Document getStylesXml() throws Exception {
        if (this.styles_xml == null)
            initStylesXml();
        return this.styles_xml.writeXml();
    }

    private Styles getStyles() throws Exception {
        if (this.styles_xml == null) {
            initStylesXml();
        }
        return this.styles_xml;
    }

    public void addStyle(Style s) throws Exception {
        getStyles().addStyle(s);
    }

    private Document getWebSettingsXml() throws Exception {
        if (this.webSettings_xml == null)
            initWebSettingsXml();
        return this.webSettings_xml.writeXml();
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
        return this.theme_xml.writeXml();
    }

    private Document getRelsXml() throws Exception {
        if (this.rels_xml == null)
            initRelsXml();
        return this.rels_xml.writeXml();
    }

    private Document getAppXml() throws Exception {
        if (this.app_xml == null)
            initAppXml();
        return this.app_xml.writeXml();
    }

    private Document getCoreXml() throws Exception {
        if (this.core_xml == null)
            initCoreXml();
        return this.core_xml.writeXml();
    }

    private void initDocumentXml() throws Exception {
        if (this.document_xml == null)
            this.document_xml = new DocumentXml();
    }

    private void initFontTableXml() throws Exception {
        if (this.fontTable_xml == null)
            this.fontTable_xml = new FontTable();
    }

    private void initSettingsXml() throws Exception {
        if (this.settings_xml == null)
            this.settings_xml = new Settings();
    }

    private void initStylesXml() throws Exception {
        if (this.ContentTypes_xml == null)
            initContentTypes();
        if (this.document_xml_rels == null)
            initDocumentXmlRels();
        this.ContentTypes_xml.addOverride("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
        this.document_xml_rels.addRelationship("styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        if (this.styles_xml == null)
            this.styles_xml = new Styles();
    }

    private void initWebSettingsXml() throws Exception {
        if (this.webSettings_xml == null)
            this.webSettings_xml = new WebSettings();
    }

    private void initContentTypes() throws Exception {
        if (this.ContentTypes_xml == null)
            this.ContentTypes_xml = new ContentTypes();
    }

    private void initDocumentXmlRels() throws Exception {
        if (this.document_xml_rels == null)
            this.document_xml_rels = new DocumentXmlRels();
    }

    private void initThemeXml() throws Exception {
        if (this.theme_xml == null)
            this.theme_xml = new Theme();
    }

    private void initRelsXml() throws Exception {
        if (this.rels_xml == null)
            this.rels_xml = new RelsRels();
    }

    private void initAppXml() throws Exception {
        if (this.app_xml == null)
            this.app_xml = new App();
    }

    private void initCoreXml() throws Exception {
        if (this.core_xml == null)
            this.core_xml = new Core();
    }

    private void saveDebug() throws Exception {
        addEntry(zos, getDocumentXml(), "word/document.xml");
        addEntry(zos, getContentTypesXml(), "[Content_Types].xml");
        addEntry(zos, getRelsXml(), "_rels/.rels");
        if (this.document_xml_rels != null) {
            addEntry(zos, getDocumentXmlRelsXml(), "word/_rels/document.xml.rels");
        }
        // if (this.styles_xml != null) {
        addEntry(zos, getStylesXml(), "word/styles.xml");
        // }
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
            Docx doc = new Docx("c:\\temp\\test.docx");
            doc.save();
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
