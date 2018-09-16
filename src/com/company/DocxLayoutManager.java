package com.company;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.util.List;
import java.util.logging.Logger;

public class DocxLayoutManager {
    //Constants
    private static final String DOCX4J_CLASS_R = "org.docx4j.wml.R";

    //Variables
    private WordprocessingMLPackage wordPackage;
    private MainDocumentPart docPart;
    private Document document;
    private Body docBody;
    private List<Object> docBodyList;
    private List<Object> paraList;
    private Logger logger = Logger.getLogger(DocxLayoutManager.class.getName());

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
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Iterate through document to get each runtext and their properties
     */
    public void docIterator() {
        if (docBodyList != null) {
            //iterate to get paragraphs
            for (int i = 0; i < docBodyList.size(); i++) {
                Object paragraph = docBodyList.get(i);
                paraList = TraversalUtil.getChildrenImpl(paragraph);

                if (paraList != null) {
                    //iterate to get runs
                    for (int j = 0; j < paraList.size(); j++) {
                        Object run = paraList.get(j);
                        String runText = getRunText(run);
                        String runPropText = getRunProp(run);

                        //Example output
                        if (runText != null) {
                            logger.info("********************\n" + "Paragraph: " + i + " Run: " + j + "\n" +
                                    runText + "\n" + runPropText + "********************");
                        }
                    }
                }
            }
        }
    }

    /**
     * Returns the text of the run
     * @param run
     */
    private String getRunText (Object run) {
        Object runContent = (TraversalUtil.getChildrenImpl(run) == null) ? null : TraversalUtil.getChildrenImpl(run).get(0);
        return runContent != null ? ((Text) ((JAXBElement) runContent).getValue()).getValue() : null;
    }

    /**
     * Returns String containing all props of run
     * @param run
     */
    private String getRunProp (Object run) {
        String props = "";
        String divider = "\n";
        RPr rpr;

        if (run != null && run.getClass().getName().equalsIgnoreCase(DOCX4J_CLASS_R)) {
            rpr = ((R) XmlUtils.unwrap(run)).getRPr();
            if (rpr != null) {
                //Font
                if (rpr.getRFonts() != null && rpr.getRFonts().getAscii() != null) {
                    props += "Font: " + rpr.getRFonts().getAscii() + divider;
                }

                //Bold
                if (rpr.getB() != null) {
                    props += "Bold" + divider;
                }

                //Italics
                if (rpr.getI()!=null) {
                    props += "Italics" + divider;
                }

                //Caps
                if (rpr.getCaps()!=null) {
                    props += "Caps" + divider;
                }

                //Small Caps
                if (rpr.getSmallCaps()!=null) {
                    props += "Small Caps" + divider;
                }

                //Strikethrough
                if (rpr.getStrike()!=null) {
                    props += "Strike" + divider;
                }

                //Font Colour
                if (rpr.getColor()!=null ) {
                    props += "Font Colour: " + rpr.getColor().getVal() + divider;
                }

                //Size
                if (rpr.getSz()!=null ) {
                    props += "Font Size: " + rpr.getSz().getVal() + divider;
                }

                //Underline;
                if (rpr.getU()!=null ) {
                    props += "Underline" + divider;
                }

                //Text Border;
                if (rpr.getBdr()!=null ) {
                    props += "Border" + divider;
                }

                //Shading;
                if (rpr.getShd()!=null ) {
                    props += "Shading" + divider;
                }
            }
        }
        return props;
    }
}
