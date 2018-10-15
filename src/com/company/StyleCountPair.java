package com.company;

public class StyleCountPair {
    private Run run;
    private int charCount;

    public StyleCountPair(Run run) {
        this.run = run;
        charCount = run.getCharCount();
    }

    public Run getRun() {
        return run;
    }

    public void setRun(Run run) {
        this.run = run;
    }

    public int getCharCount() {
        return charCount;
    }

    public void setCharCount(int charCount) {
        this.charCount = charCount;
    }
}
