package com.company;

import org.apache.log4j.BasicConfigurator;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.model.PropertyResolver;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.swing.*;
import javax.xml.bind.JAXBElement;
import java.io.FileInputStream;
import java.util.List;
import java.util.Map;

/**
 *
 * http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/index.html
 *
 * https://www.experts-exchange.com/questions/29054970/How-to-get-italic-text-using-docx4j-in-java.html
 * https://github.com/plutext/docx4j/blob/master/src/main/java/org/docx4j/model/PropertyResolver.java
 *
 * https://www.docx4java.org/forums/docx-java-f6/check-styles-in-docx-using-docx4j-t2378.html
 *
 * **/
public class Main {

    public static void main(String[] args) {
        /*BasicConfigurator.configure();

        //JFileChooser jfc = new JFileChooser();
        //int returnValue = jfc.showOpenDialog(null);
        try {
            //if (returnValue == JFileChooser.APPROVE_OPTION) {
                //wordMLPackage = WordprocessingMLPackage.load(new FileInputStream(jfc.getSelectedFile()));

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("C:/Users/Brian/Desktop/Capstone Test Doc.docx"));
            MainDocumentPart docPart = wordMLPackage.getMainDocumentPart();
            Document wmlDocEle = docPart.getJaxbElement();
            Body body = wmlDocEle.getBody();
            List<Object> paraList = TraversalUtil.getChildrenImpl(body);

            for (Object a : paraList) {
                List<Object> paraWRList = TraversalUtil.getChildrenImpl(a);
                if (paraWRList != null) {
                    for (Object b : paraWRList) {
                        if (b != null) {
                            if (b.getClass().getName().equalsIgnoreCase("org.docx4j.wml.R")) {
                                if (((R) XmlUtils.unwrap(b)).getRPr() != null) {
                                    styleCheck(a.toString(),((R) XmlUtils.unwrap(b)).getRPr());
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }*/

/*
         - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        try {

            BasicConfigurator.configure();
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("C:/Users/Brian/Desktop/Capstone Test Doc.docx"));
            MainDocumentPart docPart = wordMLPackage.getMainDocumentPart();

            FormattingLister fl = new FormattingLister();
            fl.propertyResolver = wordMLPackage.getMainDocumentPart().getPropertyResolver();

            if (wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart() == null) {
                System.out.println("no styles part!");
            } else {
                new TraversalUtil(wordMLPackage.getMainDocumentPart().getJaxbElement(), fl);

                for (Map.Entry<PartName, Part> entry : wordMLPackage.getParts().getParts().entrySet()) {

                    Part p = entry.getValue();

                    if (p instanceof HeaderPart) {
                        new TraversalUtil(((HeaderPart) p).getJaxbElement().getEGBlockLevelElts(), fl);
                    }

                    if (p instanceof FooterPart) {
                        new TraversalUtil(((FooterPart) p).getJaxbElement().getEGBlockLevelElts(), fl);
                    }
                }
            }


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
        - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            */




        try {
            BasicConfigurator.configure();
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("C:/Users/Brian/Desktop/Capstone Test Doc.docx"));
            MainDocumentPart docPart = wordMLPackage.getMainDocumentPart();
            Document wmlDocEle = docPart.getJaxbElement();
            Body body = wmlDocEle.getBody();
            List<Object> paraList = TraversalUtil.getChildrenImpl(body);

            PropertyResolver propRes;
            PPr pprDir;


            System.out.println("Number of Paragraphs = " + paraList.size());

            //iterate through paragraphs
            for (int i = 0; i < paraList.size(); i++) {
                Object a = paraList.get(i);
                List<Object> paraWRList = TraversalUtil.getChildrenImpl(a);

                //iterate through runs
                if (paraWRList != null) {
                    System.out.println("Number of Runs = " + paraWRList.size());

                    for (int j = 0; j < paraWRList.size(); j++) {
                        System.out.println("Start**************************************************");
                        //System.out.println("Run Number: " + (j + 1));
                        Object b = paraWRList.get(j);
                        Object c = null;
                        String runText = null;

                        if (TraversalUtil.getChildrenImpl(b) != null) {
                            c = TraversalUtil.getChildrenImpl(b).get(0);
                        }
                        if (c != null) {
                            runText = ((Text) ((JAXBElement) c).getValue()).getValue();
                            System.out.println("Run Text [" + (j + 1) + "]: " + runText);
                        }

                        if (b != null) {
                            if (b.getClass().getName().equalsIgnoreCase("org.docx4j.wml.R")) {
                                if (((R) XmlUtils.unwrap(b)).getRPr() != null) {
                                    styleCheck(runText,((R) XmlUtils.unwrap(b)).getRPr());
                                    System.out.println("**************************************************End");
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /*public static class FormattingLister extends TraversalUtil.CallbackImpl {

        PropertyResolver propertyResolver;

        PPr pPrDirect;

        @Override
        public List<Object> apply(Object o) {
            if (o instanceof P) {
                P p = (P)o;
                pPrDirect = propertyResolver.getEffectivePPr(p.getPPr());

                if (pPrDirect.getSpacing()!=null) {
                    System.out.println(XmlUtils.marshaltoString(pPrDirect.getSpacing()));
                }
            }

            if (o instanceof R) {
                R r = (R)o;
                RPr rPr = propertyResolver.getEffectiveRPr(r.getRPr(), pPrDirect);
                System.out.println(XmlUtils.marshaltoString(rPr));
            }
            return null;
        }
    }*/



    private static void styleCheck (String s, RPr rpr) {
        if (rpr != null) {

            //System.out.println("Start**************************************************");
            System.out.println(s);

            //Font
            if (rpr.getRFonts()!=null ) {
                System.out.println("Font: " + rpr.getRFonts().getAscii());
            }

            //Bold
            if (rpr.getB()!=null) {
                System.out.println("Bold");
            }

            //Italics
            if (rpr.getI()!=null) {
                System.out.println("Italics");
            }

            //Caps
            if (rpr.getCaps()!=null) {
                System.out.println("Caps");
            }

            //Small Caps
            if (rpr.getSmallCaps()!=null) {
                System.out.println("Small Caps");
            }

            //Strikethrough
            if (rpr.getStrike()!=null) {
                System.out.println("Strike");
            }

            //Font Colour
            if (rpr.getColor()!=null ) {
                System.out.println("Font Colour: " + rpr.getColor().getVal());
            }

            //Size
            if (rpr.getSz()!=null ) {
                System.out.println("Font Size: " + rpr.getSz().getVal());
            }

            //Underline;
            if (rpr.getU()!=null ) {
                System.out.println("Underline");
            }

            //Text Border;
            if (rpr.getBdr()!=null ) {
                System.out.println("Border");
            }

            //Shading;
            if (rpr.getShd()!=null ) {
                System.out.println("Shading");
            }

            //System.out.println("**************************************************End");
        }
    }
}