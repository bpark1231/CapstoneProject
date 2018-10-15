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

    public void printParaCharCount() {
        List<StyleCountPair> list = new ArrayList<>();
        int count = 0;
        for (Paragraph p : paraList) {
            System.out.println("paragraph count" + p.getCharCount());
            count += p.getCharCount();
        }
        System.out.println("total count: " + count);

        System.out.println("\n\nCharacter Style Consistency");


        //TODO
        List<StyleCountPair> styleList = new ArrayList<>();

        for (Paragraph p : paraList) {
            //
            styleList.addAll(p.getUniqueStyles());

            List<StyleCountPair> tempList = p.getUniqueStyles();
        }

        //set temp ids to runs
        for (int i = 0; i < styleList.size(); i++) {
            styleList.get(i).getRun().setId(i);
        }

        List<StyleCountPair> tempList = new ArrayList<>();

        for (int i = 0; i < styleList.size(); i++) {
            StyleCountPair sOuter = styleList.get(i);
            Run rOuter = sOuter.getRun();

            for (int j = 0; j < styleList.size(); j++) {
                StyleCountPair sInner = styleList.get(j);
                Run rInner = sInner.getRun();

                if (tempList.isEmpty()) {
                    tempList.add(sInner);
                } else if (rInner.hasSameStyle(rOuter) && rInner.getId() != rOuter.getId()) {
                    if (styleExistsInList(rInner, tempList)) {
                        int index = getIndexWithStyle(rInner, tempList);
                        StyleCountPair s = tempList.get(index);
                        tempList.get(index).getRun().setId(s.getRun().getCharCount() + rInner.getCharCount());
                    } else {
                        tempList.add(sInner);
                    }
                } else if (!rInner.hasSameStyle(rOuter) && !styleExistsInList(rInner, tempList)) {
                    tempList.add(sInner);
                }
            }
        }

        for (StyleCountPair s : tempList) {
            System.out.println("\"" + s.getRun().getRunText() + "\"\nchar count: " + s.getCharCount() + "\n");
        }
        System.out.println(tempList.size());
    }

    private boolean styleExistsInList(Run r, List<StyleCountPair> list) {
        boolean doesExist = false;
        for (StyleCountPair s : list) {
            if (r.hasSameStyle(s.getRun())) {
                doesExist = true;
            }
        }
        return doesExist;
    }

    private int getIndexWithStyle(Run r, List<StyleCountPair> list) {
        int index = -1;
        for (int i = 0; i < list.size(); i++){
            if (r.hasSameStyle(list.get(i).getRun())) {
                return i;
            }
        }
        return index;
    }

    public void testing() {
        System.out.println("************************************TEST START************************************");
        List<Run> allRuns = new ArrayList<>();
        for (Paragraph p : paraList) {
            allRuns.addAll(p.getRuns());
        }

        int i = 0;
        for (Run r : allRuns) {
            System.out.println("index: "+ i + r.toString());
            i++;
        }
        System.out.println("Same Style Check: " + allRuns.get(2).hasSameStyle(allRuns.get(5)));

        //paraList.get(1).getUniqueStyles();
        System.out.println("*************************************TEST END*************************************");
    }
}
