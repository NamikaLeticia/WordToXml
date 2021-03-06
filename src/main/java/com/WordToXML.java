package com;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.w3c.dom.Document;
import org.w3c.dom.Element;


public  class  WordToXML {
    public static void main(String[] args) {

        WordToXML u =new WordToXML();
       u.initializeXml();
       u.processDocToXml("D:\\xml\\src\\main\\resources\\debate.doc","debate.doc");
        u.generateXml("D:\\xml", "/xml/debate.xml");
    }

    static final Logger logger = Logger.getLogger(WordToXML.class);
    DocumentBuilderFactory docFactory = null;
    DocumentBuilder docBuilder = null;
    Element rootElement = null;
    Document docxml = null;
    boolean subHeaders = false;
    Element UrlElement = null;

   public void processDocxToXml(String path, String fileName) {
        try {
            File file = new File("D://xml//src//main//resources//debate.docx");
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
           OPCPackage oPackage = OPCPackage.open(fis);
            XWPFDocument xdoc = new XWPFDocument(oPackage);

            List<XWPFParagraph> paragraphList = xdoc.getParagraphs();
            for (XWPFParagraph paragraph : paragraphList) {

                String styleName = paragraph.getStyle();
                String paraText = paragraph.getParagraphText();
                String bulletsPoints = paragraph.getNumFmt();
                initializeXml();
                createXmlTags(styleName, paraText, bulletsPoints);

            }
            //fis.close();
             WordToXML d= new WordToXML();
                initializeXml();
            // write the content into XML file
            generateXml("D:\\xml","/xml/debate.xml");
            logger.info("Docx to Xml Conversion completed.");

        } catch (Exception ex) {
            logger.error("Exception while generating XML from DOCX" + ex.getMessage());
            System.exit(0);
        }
    }

   public void processDocToXml(String path, String fileName) {
        HWPFDocument doc = null;
        try {
            POIFSFileSystem fis = new POIFSFileSystem(new FileInputStream("D:\\xml\\src\\main\\resources\\debate.doc"));
            doc = new HWPFDocument(fis);

        } catch (Exception e) {
            logger.error("Unable to Read File..." + e.getMessage());
            System.exit(0);
        }
        try {
            WordExtractor we = new WordExtractor(doc);
            Range range = doc.getRange();

            initializeXml();

            String[] paragraphs = we.getParagraphText();

            for (int i = 0; i < paragraphs.length; i++) {
                org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);

                int j = 0;
                while (true) {
                    CharacterRun run = pr.getCharacterRun(j++);

                    StyleDescription style = doc.getStyleSheet().getStyleDescription(run.getStyleIndex());
                    String styleName = style.getName();
                    String paraText = run.text();
                    String bulletsPoints = null;

                createXmlTags(styleName, paraText,bulletsPoints);

                    if (run.getEndOffset() == pr.getEndOffset()) {
                        break;
                    }
                }
            }

            generateXml(path, fileName);

            logger.info("Document to Xml Conversion completed.");
        } catch (Exception ex) {
            logger.error("Exception while generating XML from DOC" + ex.getMessage());
            System.exit(0);
        }
    }

   public void initializeXml() {

        // initialize XML Document
        try {

           docFactory = DocumentBuilderFactory.newInstance();
            docBuilder = docFactory.newDocumentBuilder();
           docxml = docBuilder.newDocument();
          rootElement = docxml.createElement("Debate");
            docxml.appendChild(rootElement);
        } catch (ParserConfigurationException e) {
            logger.error("Exception while initializing XML" + e.getMessage());
        }

    }

    public void createXmlTags(String styleName, String paragraphText, String bulletsPoints) {

        // create XML Tags

        if (styleName != null && paragraphText.length() > 1) {
            if (styleName.equalsIgnoreCase("Style4")) {
                Element pragElement = docxml.createElement("TITLE");
                pragElement.appendChild(docxml.createTextNode(paragraphText.trim()));
                rootElement.appendChild(pragElement);
                subHeaders = true;
            } else if (styleName.equalsIgnoreCase("Default")) {
                Element pragElement = docxml.createElement("P");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                rootElement.appendChild(pragElement);
                subHeaders = true;
            }else if (styleName.equalsIgnoreCase("Normal")) {
                Element pragElement = docxml.createElement("P");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                rootElement.appendChild(pragElement);
                subHeaders = true;
            } else if (styleName.equalsIgnoreCase("BodyCopy") && bulletsPoints != null) {
                Element pragElement = docxml.createElement("LI");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                UrlElement.appendChild(pragElement);
                subHeaders = false;
            } else if (styleName.equalsIgnoreCase("BodyCopy")) {
                Element pragElement = docxml.createElement("PS");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                rootElement.appendChild(pragElement);
                subHeaders = true;
            } else if (styleName.equalsIgnoreCase("ListParagraph")) {
                Element pragElement = docxml.createElement("LI");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                UrlElement.appendChild(pragElement);
                subHeaders = false;
            } else if (styleName.equalsIgnoreCase("Subhead1")) {
                UrlElement = docxml.createElement("UL");

                Element pragElement = docxml.createElement("LI");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                UrlElement.appendChild(pragElement);
                rootElement.appendChild(UrlElement);
                subHeaders = false;

            } else {
                Element pragElement = docxml.createElement("PS");
                pragElement.appendChild(docxml.createTextNode(paragraphText));
                rootElement.appendChild(pragElement);
                subHeaders = true;
            }

        } /**else if (paragraphText.trim().length() > 1) {
            Element pragElement = docxml.createElement("P");
            pragElement.appendChild(docxml.createTextNode(paragraphText));
            rootElement.appendChild(pragElement);
            subHeaders = true;
        }*/
        else if (paragraphText.trim().isEmpty()) {
            Element pragElement = docxml.createElement("NEWLINE");
            pragElement.appendChild(docxml.createTextNode(""));
            rootElement.appendChild(pragElement);
        }

        /**else/**(subHeaders) {
            Element pragElement = docxml.createElement("NEWLINE");
            pragElement.appendChild(docxml.createTextNode(""));
            rootElement.appendChild(pragElement);
        }*/

    }

   public  void generateXml(String path, String fileName) {
        try {
            // write the content into xml file
            Transformer transformer;
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            transformer = transformerFactory.newTransformer();
           transformer.setOutputProperty(OutputKeys.METHOD, "xml");
           transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
           transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
          transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
           StreamResult result = new StreamResult(new File("D:/xml/debate.xml")  );
            DOMSource source = new DOMSource(docxml);
            transformer.transform( source, result );
        } catch (Exception e) {
            logger.error("Exception while generating XML" + e.getMessage());
        }
    }
}
