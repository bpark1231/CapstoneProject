package com.company;

import org.docx4j.TraversalUtil;

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
        if (runObjList != null) {
            for (Object r : runObjList) {
                Run run = new Run(r);
                if (run.getRunText() != null && !run.getRunText().isEmpty()) {
                    runs.add(run);
                }
            }
        }
    }

    public List<Run> getRuns() {
        return runs;
    }
}
