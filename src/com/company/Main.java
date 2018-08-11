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
 * https://www.experts-exchange.com/questions/29054970/How-to-get-italic-text-using-docx4j-in-java.html
 *
 * **/
public class Main {

    public static void main(String[] args) {
	// write your code here
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
            //List<Object> paraWRList = TraversalUtil.getChildrenImpl(paraList.get(1));

            for (Object a : paraList) {
                List<Object> paraWRList = TraversalUtil.getChildrenImpl(a);
                for (Object b : paraWRList)
                    if (b != null) {
                        if (((R) XmlUtils.unwrap(b)).getRPr() != null) {
                            if (((R) XmlUtils.unwrap(b)).getRPr().getB() != null) {
                                //System.out.println( + ": bold");
                                System.out.println(a.toString() + ": bold");
                            } else {
                                System.out.println(a.toString() + ": not bold");
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
}