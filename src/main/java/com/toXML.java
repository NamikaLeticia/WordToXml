package com;

import org.apache.log4j.Logger;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;

public class toXML {
    public static void main(String[] args) {

        toXML u =new toXML();
        //u.initializeXml();
      //  u.processDocToXml("D:\\xml\\src\\main\\resources\\debate.doc","debate.doc");
       u.generateXml("D:\\xml", "/xml/debate.xml");
    }

    static final Logger logger = Logger.getLogger(toXML.class);
    DocumentBuilderFactory docFactory = null;
    DocumentBuilder docBuilder = null;
    Element rootElement = null;
    Document docxml = null;
    boolean subHeaders = false;
    Element UrlElement = null;
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
    public  void generateXml(String path, String fileName) {
        try {
            // write the content into xml file
            Transformer transformer;
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.METHOD, "xml");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
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
