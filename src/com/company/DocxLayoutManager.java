package com.company;

import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.util.ArrayList;
import java.util.List;

public class DocxLayoutManager {
    //Variables
    private WordprocessingMLPackage wordPackage;
    private MainDocumentPart docPart;
    private Document document;
    private Body docBody;
    private List<Object> docBodyList;
    private List<Paragraph> paraList = new ArrayList<>();

    /**
     * Constructor
     * @param docPath path of document as String
     */
    public DocxLayoutManager(String docPath) {
        try {
            wordPackage = WordprocessingMLPackage.load(new java.io.File(docPath));
            docPart = wordPackage.getMainDocumentPart();
            document = docPart.getJaxbElement();
            docBody = document.getBody();
            docBodyList = TraversalUtil.getChildrenImpl(docBody);
            populateParaList();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Iterate through document to get each runtext and their properties
     */
    private void populateParaList() {
        if (docBodyList != null) {
            //iterate to get paragraphs
            for (Object paragraph : docBodyList) {
                Paragraph p = new Paragraph(paragraph);
                paraList.add(p);
            }
        }
    }

    public String toString() {
        String s = "";
        if (paraList != null) {
            for (Paragraph p : paraList) {
                for (Run r : p.getRuns()) {
                    s += r.toString();
                }
            }
        }
        return s;
    }
}
