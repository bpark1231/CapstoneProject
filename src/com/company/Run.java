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
    private static final String DEFAULT_FONT_VALUE = "default";
    private static final int INT_EMPTY_VALUE = 0;

    //Variables
    private String runText;
    private String font;
    private BigInteger fontSize;
    private int id;
    private int fontColour;
    private int charCount;
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
        font = (rpr != null && rpr.getRFonts() != null && rpr.getRFonts().getAscii() != null) ? rpr.getRFonts().getAscii() : DEFAULT_FONT_VALUE;
        fontColour = (rpr != null && rpr.getColor() != null) ? Integer.parseInt(rpr.getColor().getVal()) : INT_EMPTY_VALUE;
        fontSize = (rpr != null && rpr.getSz() != null) ? rpr.getSz().getVal() : BigInteger.valueOf(0);
        bold = rpr != null && rpr.getB() != null;
        italics = rpr != null && rpr.getI() != null;
        caps = rpr != null && rpr.getCaps() != null;
        smallCaps = rpr != null && rpr.getSmallCaps() != null;
        strikeThrough = rpr != null && rpr.getStrike() != null;
        underline = rpr != null && rpr.getU() != null;
        border = rpr != null && rpr.getBdr() != null;
        shading = rpr != null && rpr.getShd() != null;
        charCount = runText.length();
    }

    private String extractRunText (Object run) {
        Object runContent = (TraversalUtil.getChildrenImpl(run) == null) ? null : TraversalUtil.getChildrenImpl(run).get(0);
        return runContent != null ? ((Text) ((JAXBElement) runContent).getValue()).getValue() : null;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
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

    public int getCharCount() {
        return runText != null && !runText.equals("") ? runText.length() : 0;
    }

    public boolean hasSameStyle(Run r) {

        if (font != null && r.getFont() != null) {
            return (font.equals(r.getFont()) && fontColour == r.getFontColour() && fontSize.equals(r.getFontSize()) && bold == r.isBold()
                    && underline == r.isUnderline() && italics == r.isItalics() && smallCaps == r.isSmallCaps() && caps == r.isCaps()
                    && strikeThrough == r.isStrikeThrough() && border == r.isBorder() && shading == r.isShading());
        }
        return false;
    }

    public boolean isEquals(Run r) {
        return r.getRunText().equals(runText) && hasSameStyle(r);
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
