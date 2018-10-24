package com.company;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.wml.P;

import java.util.ArrayList;
import java.util.List;

/**todo
 * total char count + char count of differently styled text to check for consistency
 * total count of different combination of text styles
 * line spacing consistency
 * have certain tests that check font/para length/etc
 * make sure output can differentiate different paragraphs
 */

public class Paragraph {
    private List<Run> runs = new ArrayList<>();

    public Paragraph(Object para) {
        List<Object> runObjList;
        runObjList = para != null ? TraversalUtil.getChildrenImpl(para) : null;
        String heading = Globals.DEFAULT_HEADING_TEXT;
        if (para != null
                && ((P) XmlUtils.unwrap(para)).getPPr() != null
                && ((P) XmlUtils.unwrap(para)).getPPr().getPStyle() != null
                && ((P) XmlUtils.unwrap(para)).getPPr().getPStyle().getVal() != null) {
            heading = ((P) XmlUtils.unwrap(para)).getPPr().getPStyle().getVal();
        }
        if (runObjList != null) {
            for (Object r : runObjList) {
                Run run = new Run(r);
                if (run.getRunText()!= null && !run.getRunText().equals("")) {
                    run.setHeading(heading);
                    runs.add(run);
                }
            }
        }
    }

    /**
     * @return a list of all the runs in the paragraph
     */
    public List<Run> getRuns() {
        return runs;
    }

    /**
     * return the char count for the paragraph
     * @return
     */
    public int getCharCount() {
        int count = 0;
        for (Run run : runs) {
            count += run.getCharCount();
        }
        return count;
    }

    /**
     * return a list of the unique styles in the paragraph
     */
    public List<StyleCountPair> getUniqueStyles() {
        List<StyleCountPair> styleList = new ArrayList<>();

        for (Run r : runs) {
            if (styleList.size() != 0) {
                for (int i = 0; i < styleList.size(); i++) {
                    StyleCountPair s = styleList.get(i);
                    if (!r.hasSameStyle(s.getRun()) && !styleExistsInList(r, styleList)) {
                        styleList.add(new StyleCountPair(r));
                    }
                    else if (r.hasSameStyle(s.getRun())){
                        s.setCharCount(r.getCharCount() + s.getRun().getCharCount());
                    }
                }
            } else {
                styleList.add(new StyleCountPair(r));
            }
        }

        return styleList;
    }

    /**
     * check to see if the run exists within given list
     */
    private boolean styleExistsInList(Run r, List<StyleCountPair> list) {
        boolean doesExist = false;
        for (StyleCountPair s : list) {
            doesExist = r.hasSameStyle(s.getRun());
        }
        return doesExist;
    }
}
