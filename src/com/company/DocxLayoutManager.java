package com.company;

import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.math.BigInteger;
import java.text.DecimalFormat;
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
    private List<StyleCountPair> stylesList = new ArrayList<>();

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
            populateUniqueStyleList();
        }
    }

    /**
     * Display test results from different tests
     * NOTE: there is currently only 1 test, however this will ensure adding tests later is easier
     */
    public void displayTestInfo() {
       testUniqueStyleConsistency();
    }

    /**
     * TEST: compares each unique style to total char count to test style consistency
     */
    private void testUniqueStyleConsistency() {
        int totalCharCount = 0;
        DecimalFormat df = new DecimalFormat(".00");
        StyleCountPair mostCommon = new StyleCountPair(null);
        for (StyleCountPair s : stylesList) {
            totalCharCount += s.getCharCount();
        }

        System.out.println(Globals.STRING_DIVIDER_ENTER + "Style Consistency: ");

        for (int i = 0; i <  stylesList.size(); i++) {
            StyleCountPair s = stylesList.get(i);
            int charcount = s.getCharCount();
            double percent = (double)charcount / (double)totalCharCount * 100;

            String string = "Style " + (i + 1) + ": ";
            string += charcount + "/" + totalCharCount;
            string += " (" + df.format(percent) +"%)";
            System.out.println(string);

            if (s.getCharCount() > mostCommon.getCharCount()) {
                mostCommon = s;
            }
        }

        double percent = (double)mostCommon.getCharCount() / (double)totalCharCount * 100;
        String testMsg = "\nMost common style:\n" + mostCommon.getRun().toString(false);
        testMsg += "Consistency: " + df.format(percent);
        System.out.println(testMsg);
        if (percent < 67) {
            System.out.println("This document does not have a consistent style for the body. Consider using less styles.");
        }
    }

    /**
     * prints the character count for each paragraph
     * and the total character count
     */
    public void printParaCharCount() {
        int count = 0;
        for (Paragraph p : paraList) {
            System.out.println("paragraph count" + p.getCharCount());
            count += p.getCharCount();
        }
        System.out.println("total count: " + count);
    }

    /**
     * populates the styleslist field with unique styles that exist within document
     * NOTE: char counter is not working properly
     */
    private void populateUniqueStyleList() {
        List<StyleCountPair> styleList = new ArrayList<>();

        for (Paragraph p : paraList) {
            styleList.addAll(p.getUniqueStyles());
        }

        //set temp ids to runs
        for (int i = 0; i < styleList.size(); i++) {
            styleList.get(i).getRun().setId(i);
        }

        for (int i = 0; i < styleList.size(); i++) {
            StyleCountPair sOuter = styleList.get(i);
            Run rOuter = sOuter.getRun();

            for (int j = 0; j < styleList.size(); j++) {
                StyleCountPair sInner = styleList.get(j);
                Run rInner = sInner.getRun();

                if (stylesList.isEmpty()) {
                    stylesList.add(sInner);
                } else if (rInner.hasSameStyle(rOuter) && rInner.getId() != rOuter.getId()) {
                    if (styleExistsInList(rInner, stylesList)) {
                        int index = getIndexWithStyle(rInner, stylesList);
                        StyleCountPair s = stylesList.get(index);
                        if (rInner.getId() != s.getRun().getId()) {
                            s.setCharCount(s.getCharCount() + rInner.getCharCount());
                        }
                    } else {
                        stylesList.add(sInner);
                    }
                } else if (!rInner.hasSameStyle(rOuter) && !styleExistsInList(rInner, stylesList)) {
                    stylesList.add(sInner);
                }
            }
        }
    }

    /**
     * prints the unique styles that exist within the document
     */
    public void printUniqueStyles() {
        for (int i = 0; i < stylesList.size(); i++) {
            StyleCountPair s = stylesList.get(i);
            Run r = s.getRun();
            String styleString = "";

            styleString += "Style " + (i + 1) + ":" + Globals.STRING_DIVIDER_ENTER;
            styleString += "Character Count: " + s.getCharCount() + Globals.STRING_DIVIDER_ENTER;

            //print style only if not default/empty
            styleString += !r.getFont().equals(Globals.DEFAULT_FONT_VALUE) ? "Font: " + r.getFont() + Globals.STRING_DIVIDER_ENTER : "";
            styleString += !r.getHeading().equalsIgnoreCase(Globals.DEFAULT_HEADING_TEXT) ? r.getHeading() + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.getFontColour() != Globals.INT_EMPTY_VALUE ? "Font Colour: " + r.getFontColour() + Globals.STRING_DIVIDER_ENTER : "";
            styleString += !r.getFontSize().equals(BigInteger.valueOf(Globals.INT_EMPTY_VALUE)) ? "Font Size: " + r.getFontSize() + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isBold() ? Globals.RUN_PROPERTY_BOLD + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isItalics() ? Globals.RUN_PROPERTY_ITALICS + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isCaps() ? Globals.RUN_PROPERTY_CAPS + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isSmallCaps() ? Globals.RUN_PROPERTY_SMALL_CAPS + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isStrikeThrough() ? Globals.RUN_PROPERTY_STRIKETHROUGH + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isUnderline() ? Globals.RUN_PROPERTY_UNDERLINE + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isBorder() ? Globals.RUN_PROPERTY_BORDER + Globals.STRING_DIVIDER_ENTER : "";
            styleString += r.isShading() ? Globals.RUN_PROPERTY_SHADING + Globals.STRING_DIVIDER_ENTER : "";

            System.out.println(styleString);
            //System.out.println("\"" + s.getRun().getRunText() + "\"\nchar count: " + s.getCharCount() + Globals.STRING_DIVIDER_ENTER);
        }
    }

    /**
     * checks if given run exists within given list
     * @param r Run object that you wish to check if exists in list
     * @param list List array that you wish to see if run exists
     */
    private boolean styleExistsInList(Run r, List<StyleCountPair> list) {
        boolean doesExist = false;
        for (StyleCountPair s : list) {
            if (r.hasSameStyle(s.getRun())) {
                doesExist = true;
            }
        }
        return doesExist;
    }

    /**
     * returns the index of list if the list contains an item with the same style
     * @param r run object that you wish to check if has same style in list
     * @param list list of style that you wish to check run against
     */
    private int getIndexWithStyle(Run r, List<StyleCountPair> list) {
        int index = -1;
        for (int i = 0; i < list.size(); i++){
            if (r.hasSameStyle(list.get(i).getRun())) {
                return i;
            }
        }
        return index;
    }

    /**
     * @return string containing every run and its respective style properties
     */
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
