package com.company;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;

import javax.xml.bind.JAXBElement;
import java.math.BigInteger;

public class Run {
    //Constants
    private static final String DOCX4J_CLASS_R = "org.docx4j.wml.R";
    private static final int INT_EMPTY_VALUE = 0;

    //Variables
    private String runText;
    private String font;
    private BigInteger fontSize;
    private int fontColour;
    private boolean bold;
    private boolean caps;
    private boolean border;
    private boolean italics;
    private boolean shading;
    private boolean underline;
    private boolean smallCaps;
    private boolean strikeThrough;

    public Run(Object run) {
        runText = extractRunText(run);
        RPr rpr;
        if (run != null && run.getClass().getName().equalsIgnoreCase(DOCX4J_CLASS_R)) {
            rpr = ((R) XmlUtils.unwrap(run)).getRPr();
            init(rpr);
        }
    }

    /**
     * Extract properties from rpr
     */
    private void init(RPr rpr) {
        font = (rpr != null && rpr.getRFonts() != null && rpr.getRFonts().getAscii() != null) ? rpr.getRFonts().getAscii() : null;
        fontColour = (rpr != null && rpr.getColor() != null) ? Integer.parseInt(rpr.getColor().getVal()) : INT_EMPTY_VALUE;
        fontSize = (rpr != null && rpr.getSz() != null) ? rpr.getSz().getVal() : null;
        bold = rpr != null && rpr.getB() != null;
        italics = rpr != null && rpr.getI() != null;
        caps = rpr != null && rpr.getCaps() != null;
        smallCaps = rpr != null && rpr.getSmallCaps() != null;
        strikeThrough = rpr != null && rpr.getStrike() != null;
        underline = rpr != null && rpr.getU() != null;
        border = rpr != null && rpr.getBdr() != null;
        shading = rpr != null && rpr.getShd() != null;
    }

    private String extractRunText (Object run) {
        Object runContent = (TraversalUtil.getChildrenImpl(run) == null) ? null : TraversalUtil.getChildrenImpl(run).get(0);
        return runContent != null ? ((Text) ((JAXBElement) runContent).getValue()).getValue() : null;
    }

    public String getRunText() {
        return runText;
    }

    public String getFont() {
        return font;
    }

    public BigInteger getFontSize() {
        return fontSize;
    }

    public int getFontColour() {
        return fontColour;
    }

    public boolean isBold() {
        return bold;
    }

    public boolean isCaps() {
        return caps;
    }

    public boolean isBorder() {
        return border;
    }

    public boolean isItalics() {
        return italics;
    }

    public boolean isShading() {
        return shading;
    }

    public boolean isUnderline() {
        return underline;
    }

    public boolean isSmallCaps() {
        return smallCaps;
    }

    public boolean isStrikeThrough() {
        return strikeThrough;
    }

    public String toString() {
        String s = "";
        String divider = "\n";

        s += divider + "\"" + runText + "\"" + divider;
        s += "Font: " + font + divider;
        s += "Font Colour: " + fontColour + divider;
        s += "Font Size: " + fontSize + divider;
        s += "Bold: " + bold + divider;
        s += "Underline: " + underline + divider;
        s += "Italics: " + italics + divider;
        s += "Caps: " + caps + divider;
        s += "Small Caps: " + smallCaps + divider;
        s += "Strike: " + strikeThrough + divider;
        s += "Border: " + border + divider;
        s += "Shading: " + shading + divider;

        return s;
    }
}
