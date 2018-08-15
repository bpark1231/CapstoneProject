package com.company;

import org.apache.log4j.BasicConfigurator;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.swing.*;
import java.io.FileInputStream;
import java.util.List;

/**
 *
 * http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/index.html
 *
 * https://www.experts-exchange.com/questions/29054970/How-to-get-italic-text-using-docx4j-in-java.html
 * https://github.com/plutext/docx4j/blob/master/src/main/java/org/docx4j/model/PropertyResolver.java
 *
 * **/
public class Main {

    public static void main(String[] args) {
        BasicConfigurator.configure();

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
                            /*if (((R) XmlUtils.unwrap(b)).getRPr().getB() != null) {
                                //System.out.println( + ": bold");
                                System.out.println(a.toString() + ": bold");
                            } else {
                                System.out.println(a.toString() + ": not bold");
                            }*/
                                    styleCheck(a.toString(),((R) XmlUtils.unwrap(b)).getRPr());
                                }
                            }
                        }
                    }
                }
            }
                /*for (int i = 0; i < obj.size(); i++) {
                    obj.get(i);
                    PPr ppr = ((P) XmlUtils.unwrap(obj)).getPPr();
                }*/


            //}
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void styleCheck (String s,RPr rpr) {
        if (rpr != null) {

            System.out.println("**************************************************");
            System.out.println(s);

            //Font
            if (rpr.getRFonts()!=null ) {
                System.out.println(rpr.getRFonts().toString());
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
                System.out.println("Coloured Font");
            }

            //Size
            if (rpr.getSz()!=null ) {
                System.out.println("Font Size");
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

            System.out.println("**************************************************");
        }
    }
}