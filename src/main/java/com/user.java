package com;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.io.File;

public class user {


    public static void main(String[] args) throws ParserConfigurationException,
            TransformerException {
        DocumentBuilder builder = null;
        try {
            builder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        } catch (ParserConfigurationException e) {
            throw e;
        }
        Document document = builder.newDocument();
        // create root element
        Element root = document.createElement("Users");
        // attach it to the document
        document.appendChild(root);
        // create user node
        Element user = document.createElement("User");
        // create its id attribute
        user.setAttribute("id", "2");
        // add user node to root node
        root.appendChild(user);
        // create name node and set its value
        Element userName = document.createElement("name");
        userName.setTextContent("codippa");
        // attach this node to user node
        user.appendChild(userName);
        // write xml
        Transformer transformer;
        try {
            TransformerFactory transformerFactory = TransformerFactory
                    .newInstance();
            transformer = transformerFactory.newTransformer();
            Result output = new StreamResult(new File("/xml/codippa.xml"));
            Source input = new DOMSource(document);
            // if you want xml to be properly formatted
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.transform(input, output);
        } catch (TransformerConfigurationException e) {
            throw e;
        } catch (TransformerException e) {
            throw e;
        }
    }
}
