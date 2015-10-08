package com.filipenunes.generatorxls;

import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.BoldStyle;
import jxl.write.WritableFont;

/**
 * Created by Filipe Nunes on 08/10/2015.
 */
public class StylesCell {

    private final WritableFont fontText;
    private final Colour colourText;
    private final boolean italicText;
    private final boolean boldText;

    private final Alignment alignmentCell;
    private final Colour backgroudCell;


    public StylesCell(WritableFont fontText, Colour colourText, boolean italicText, boolean boldText, Alignment alignmentCell, Colour backgroudCell) {
        this.fontText = fontText;
        this.colourText = colourText;
        this.italicText = italicText;
        this.boldText = boldText;
        this.alignmentCell = alignmentCell;
        this.backgroudCell = backgroudCell;
    }

    public WritableFont getFontText() {
        return fontText;
    }

    public Colour getColourText() {
        return colourText;
    }

    public boolean isItalicText() {
        return italicText;
    }

    public boolean   getBoldText() {
        return boldText;
    }

    public Alignment getAlignmentCell() {
        return alignmentCell;
    }

    public Colour getBackgroudCell() {
        return backgroudCell;
    }
}
