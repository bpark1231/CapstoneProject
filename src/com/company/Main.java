package com.company;

import org.apache.log4j.BasicConfigurator;

import java.util.Scanner;

import static javafx.application.Platform.exit;

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

    private static final int MENU_EXIT_CODE = -1;

    public static void main(String[] args) {
        BasicConfigurator.configure();
        DocxLayoutManager docManager = new DocxLayoutManager("C:/Users/Brian/Desktop/Capstone Test Doc.docx");
        //DocxLayoutManager docManager = new DocxLayoutManager("C:/Users/Brian/Desktop/2018_Resume_Brian Park_DigiZoo.docx");
        docManager.displayTestInfo();

        int option = 0;
        while (option != MENU_EXIT_CODE) {
            option = menu();

            switch (option) {
                case 1:
                    System.out.println(docManager.toString());
                    break;
                case 2:
                    docManager.printUniqueStyles();
                    break;
                case 3:
                    option = MENU_EXIT_CODE;
                    break;
                default: System.out.println("Enter valid option.");
            }
        }
    }

    public static int menu() {
        Scanner userKey = new Scanner(System.in);
        System.out.println(Globals.STRING_DIVIDER_ENTER + "Menu");
        System.out.println("1. Print All Runs");
        System.out.println("2. Print All Unique Styles");
        System.out.println("3. Exit");

        if (userKey.hasNextInt()) {
            return userKey.nextInt();
        }
        return 0;
    }
}