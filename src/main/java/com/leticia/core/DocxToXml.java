package com.leticia.core;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.zwobble.mammoth.DocumentConverter;
import org.zwobble.mammoth.Result;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.util.Set;

public class DocxToXml {

    public static String convertDocX2Xml(String docxFilePath, String docxFileName, String xmlDestinationPath) {
    String resultPath = "";
        try {

            DocumentConverter converter = new DocumentConverter();
            Result<String> result = converter.convertToHtml(new File(docxFilePath));
            String html = result.getValue(); // The generated HTML
            Set<String> warnings = result.getWarnings(); // Any warnings during conversion

            String xmlWithRoot = addRootElementToBeWellFormed(html);
            Document doc = convertStringToXMLDocument(xmlWithRoot);

            resultPath = generateXmlFile(doc, xmlDestinationPath.concat("/").concat(docxFileName).concat(".xml"));

        } catch (Exception e) {
            e.printStackTrace();
        }
        return resultPath;
    }

    private static String addRootElementToBeWellFormed(String XmlStr) {
        String root = "";
        String rootElementOpeningTag = "<root>";
        root = rootElementOpeningTag.concat(XmlStr);
        String rootElementClosingTag = "</root>";
        root = root.concat(rootElementClosingTag);
        System.out.println("root:" + root);
        return root;
    }

    private static Document convertStringToXMLDocument(String xmlString) {
        System.out.println("obtained: " + xmlString);
        //Parser that produces DOM object trees from XML content
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

        //API to obtain DOM Document instance
        DocumentBuilder builder = null;
        try {
            //Create DocumentBuilder with default configuration
            builder = factory.newDocumentBuilder();

            //Parse the content to Document object
            Document doc = builder.parse(new InputSource(new StringReader(xmlString)));

            return doc;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static String generateXmlFile(Document doc, String filePath) {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = null;
        try {
            transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            StreamResult console = new StreamResult(System.out);
            StreamResult file = new StreamResult(new File(filePath));

            // write data
            transformer.transform(new DOMSource(doc), file);
            transformer.transform(new DOMSource(doc), console);

        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        }
        return filePath;
    }

}