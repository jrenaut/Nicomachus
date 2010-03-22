package docx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public abstract class DocxXml {
    protected Document xml;
    protected Element root;

    protected abstract void initXml() throws Exception;

    protected abstract Document writeXml();
}
