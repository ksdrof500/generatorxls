package com.filipenunes.generatorxls;

import android.os.Environment;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.demo.Formulas;
import jxl.format.BoldStyle;
import jxl.format.Colour;
import jxl.read.biff.Formula;
import jxl.write.Alignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * Created by Filipe Nunes on 07/10/2015.
 */
public class GeneratorXLS {


    /**
    *
    * @param fileName - the name to give the new workbook file
    * @return - a new WritableWorkbook with the given fileName
    */
    public WritableWorkbook createWorkbook(String fileName){
        //exports must use a temp file while writing to avoid memory hogging
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setUseTemporaryFileDuringWrite(true);
        wbSettings.setLocale(new Locale("pt", "BR"));

        //get the sdcard's directory
        File sdCard = Environment.getExternalStorageDirectory();
        //add on the your app's path
        File dir = new File(sdCard.getAbsolutePath() + "/JExcelApiTest");
        //make them in case they're not there
        if(!dir.exists()) {
            dir.mkdirs();
        }
        //create a standard java.io.File object for the Workbook to use
        File wbfile = new File(dir,fileName);

        WritableWorkbook wb = null;

        try{
            //create a new WritableWorkbook using the java.io.File and
            //WorkbookSettings from above
            wb = Workbook.createWorkbook(wbfile, wbSettings);

        }catch(IOException ex){


        }

        return wb;
    }


    public File finishWorkbook(String name, WritableWorkbook wb) throws IOException, WriteException {

        wb.close();

        //get the sdcard's directory
        File sdCard = Environment.getExternalStorageDirectory();
        //add on the your app's path
        File dir = new File(sdCard.getAbsolutePath() + "/JExcelApiTest");
        //Create file.
        File file = new File(dir, name);

        return file;
    }

    /**
     *
     * @param wb - WritableWorkbook to create new sheet in
     * @param sheetName - name to be given to new sheet
     * @param sheetIndex - position in sheet tabs at bottom of workbook
     * @return - a new WritableSheet in given WritableWorkbook
     */
    public WritableSheet createSheet(WritableWorkbook wb,
                                     String sheetName, int sheetIndex){
        //create a new WritableSheet and return it
        return wb.createSheet(sheetName, sheetIndex);
    }

    /**
     *
     * @param columnPosition - column to place new cell in
     * @param rowPosition - row to place new cell in
     * @param contents - string value to place in cell
     * @param sheet - WritableSheet to place cell in
     * @throws RowsExceededException - thrown if adding cell exceeds .xls row limit
     * @throws WriteException - Idunno, might be thrown
     */
    public void writeCell(int columnPosition, int rowPosition, String contents,
                          WritableSheet sheet) throws RowsExceededException, WriteException {

        //create a new cell with contents at position
        Label newCell = new Label(columnPosition,rowPosition,contents);
        sheet.addCell(newCell);
    }


    public void writeCellHorizontal(int columnInitial, int rowInitial, HashMap<String, StylesCell> contents,
                          WritableSheet sheet) throws  WriteException {


        Set<String> keys = contents.keySet();
        for (String key : keys) {
            //create a new cell with contents at position
            Label newCell = new Label(columnInitial, rowInitial, key);

            //Assign Styles
            if(contents.get(key) != null) {
                WritableCellFormat cellFormat = AssingStylesCell(contents.get(key));
                newCell.setCellFormat(cellFormat);
            }

            sheet.addCell(newCell);
            columnInitial++;
            }

    }


    public void writeCellVertical(int columnInitial, int rowInitial, HashMap<String, StylesCell> contents,
                                          WritableSheet sheet) throws  WriteException {


        Set<String> keys = contents.keySet();
        for (String key : keys) {
            //create a new cell with contents at position
            Label newCell = new Label(columnInitial, rowInitial, key);

            //Assign Styles
            if(contents.get(key) != null) {
                WritableCellFormat cellFormat = AssingStylesCell(contents.get(key));
                newCell.setCellFormat(cellFormat);
            }

            sheet.addCell(newCell);
            rowInitial++;
        }

    }


    public void writeCellHorizontalStyleAll(int columnInitial, int rowInitial, List<String> contents,
                                          StylesCell styles, WritableSheet sheet) throws  WriteException {

        //Assign Styles
        WritableCellFormat cellFormat = AssingStylesCell(styles);

        for (String key : contents) {
            //create a new cell with contents at position
            Label newCell = new Label(columnInitial, rowInitial, key);
            newCell.setCellFormat(cellFormat);
            sheet.addCell(newCell);
            columnInitial++;
        }

    }


    public void writeCellVerticalStyleAll(int columnInitial, int rowInitial, List<String> contents,
                                                StylesCell styles, WritableSheet sheet) throws  WriteException {

        //Assign Styles
        WritableCellFormat cellFormat = AssingStylesCell(styles);

        for (String key : contents) {
            //create a new cell with contents at position
            Label newCell = new Label(columnInitial, rowInitial, key);
            newCell.setCellFormat(cellFormat);
            sheet.addCell(newCell);
            rowInitial++;
        }

    }

    public void writeList(int columnInitial, int rowInitial, HashMap<String, List<String>> contents,
                          StylesCell stylesHeader, StylesCell styles, WritableSheet sheet) throws WriteException {

        //Assign Styles
        WritableCellFormat cellHeaderFormat = AssingStylesCell(stylesHeader);
        WritableCellFormat cellFormat = AssingStylesCell(styles);

        Set<String> keys = contents.keySet();
        int row;
        for (String key : keys) {
            //create a new cell header
            Label newCell = new Label(columnInitial, rowInitial, key);
            newCell.setCellFormat(cellHeaderFormat);
            sheet.addCell(newCell);

            //create a new cell data
            List<String> data = contents.get(key);
            row = rowInitial;

            for(int r = 0; r < data.size(); r++){
                Label dataCell = new Label(columnInitial, row++, data.get(r));
                dataCell.setCellFormat(cellFormat);
                sheet.addCell(newCell);
            }

            columnInitial++;
        }

    }

    private WritableCellFormat AssingStylesCell(StylesCell styles) throws WriteException {


        WritableFont cellFont;

        // Font
        if(styles.getFontText() != null) {
            cellFont = new WritableFont(WritableFont.createFont(String.valueOf(styles.getFontText())), 12);
        }else{
            cellFont = new WritableFont(WritableFont.ARIAL, 12);
        }
        // Colour
        if(styles.getColourText() != null) {
            cellFont.setColour(styles.getColourText());
        }
        // Italic
        if(styles.isItalicText()) {
            cellFont.setItalic(true);
        }
        // Negrito
        if(styles.getBoldText()) {
            cellFont.setBoldStyle(WritableFont.BOLD);
        }

        WritableCellFormat cellFormat = new WritableCellFormat(cellFont);

        // Alignment
        if (styles.getAlignmentCell() != null) {
            cellFormat.setAlignment(styles.getAlignmentCell());
        }
        // Backgraund
        if (styles.getBackgroudCell() != null) {
            cellFormat.setBackground(styles.getBackgroudCell());
        }

        return cellFormat;

    }



}
