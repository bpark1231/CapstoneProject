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
    private static final String DOCX4J_CLASS_JAXBELEMENT = "javax.xml.bind.JAXBElement";
    private static final String DOCX4J_CLASS_TEXT = "org.docx4j.wml.Text";

    //Variables
    private String runText;
    private String font;
    private String heading;
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
     * https://www.docx4java.org/forums/docx-java-f6/recognizing-headings-t1268.html
     */
    private void init(RPr rpr) {
        font = (rpr != null && rpr.getRFonts() != null && rpr.getRFonts().getAscii() != null) ? rpr.getRFonts().getAscii() : Globals.DEFAULT_FONT_VALUE;
        fontColour = (rpr != null && rpr.getColor() != null) ? Integer.parseInt(rpr.getColor().getVal()) : Globals.INT_EMPTY_VALUE;
        fontSize = (rpr != null && rpr.getSz() != null) ? rpr.getSz().getVal() : BigInteger.valueOf(0);
        bold = rpr != null && rpr.getB() != null;
        italics = rpr != null && rpr.getI() != null;
        caps = rpr != null && rpr.getCaps() != null;
        smallCaps = rpr != null && rpr.getSmallCaps() != null;
        strikeThrough = rpr != null && rpr.getStrike() != null;
        underline = rpr != null && rpr.getU() != null;
        border = rpr != null && rpr.getBdr() != null;
        shading = rpr != null && rpr.getShd() != null;
        charCount = runText != null ? runText.length() : 0;
    }

    /**
     * @return the text component of a given run
     */
    private String extractRunText (Object run) {
        if (TraversalUtil.getChildrenImpl(run) != null) {
            int last = TraversalUtil.getChildrenImpl(run).size() - 1;
            Object runContent = TraversalUtil.getChildrenImpl(run).get(last);
            if (runContent != null && runContent.getClass().getName().equalsIgnoreCase(DOCX4J_CLASS_JAXBELEMENT)) {
                if (((JAXBElement) runContent).getValue().getClass().getName().equalsIgnoreCase(DOCX4J_CLASS_TEXT)) {
                    return ((Text) ((JAXBElement) runContent).getValue()).getValue();
                }
            }
        }
        return null;
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

    public String getHeading() {
        return heading;
    }

    public void setHeading(String heading) {
        this.heading = heading;
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
        return charCount;
    }

    /**
     * Checks if run has same style as given run
     */
    public boolean hasSameStyle(Run r) {

        if (font != null && r.getFont() != null) {
            return (font.equals(r.getFont()) && heading.equals(r.getHeading()) && fontColour == r.getFontColour() && fontSize.equals(r.getFontSize()) && bold == r.isBold()
                    && underline == r.isUnderline() && italics == r.isItalics() && smallCaps == r.isSmallCaps() && caps == r.isCaps()
                    && strikeThrough == r.isStrikeThrough() && border == r.isBorder() && shading == r.isShading());
        }
        return false;
    }

    /**
     * checks if run is equals to a given run
     */
    public boolean isEquals(Run r) {
        return r.getRunText().equals(runText) && hasSameStyle(r);
    }

    public String toString(boolean showText) {
        String s = "";

        if (showText) {
            s += Globals.STRING_DIVIDER_ENTER + "\"" + runText + "\"" + Globals.STRING_DIVIDER_ENTER;
        }
        s += "Font: " + font + Globals.STRING_DIVIDER_ENTER;
        s += "Heading: " + heading + Globals.STRING_DIVIDER_ENTER;
        s += "Font Colour: " + fontColour + Globals.STRING_DIVIDER_ENTER;
        s += "Font Size: " + fontSize + Globals.STRING_DIVIDER_ENTER;
        s += "Bold: " + bold + Globals.STRING_DIVIDER_ENTER;
        s += "Underline: " + underline + Globals.STRING_DIVIDER_ENTER;
        s += "Italics: " + italics + Globals.STRING_DIVIDER_ENTER;
        s += "Caps: " + caps + Globals.STRING_DIVIDER_ENTER;
        s += "Small Caps: " + smallCaps + Globals.STRING_DIVIDER_ENTER;
        s += "Strike: " + strikeThrough + Globals.STRING_DIVIDER_ENTER;
        s += "Border: " + border + Globals.STRING_DIVIDER_ENTER;
        s += "Shading: " + shading + Globals.STRING_DIVIDER_ENTER;

        return s;
    }

    public String toString() {
        return toString(true);
    }
}
