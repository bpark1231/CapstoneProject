package com.company;

import org.apache.log4j.BasicConfigurator;

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
        BasicConfigurator.configure();
        DocxLayoutManager docManager = new DocxLayoutManager("C:/Users/Brian/Desktop/Capstone Test Doc.docx");
        //DocxLayoutManager docManager = new DocxLayoutManager("C:/Users/Brian/Desktop/Logical Architecture.docx");
        System.out.println(docManager.toString());
        docManager.printParaCharCount();
        //docManager.testing();
    }
}