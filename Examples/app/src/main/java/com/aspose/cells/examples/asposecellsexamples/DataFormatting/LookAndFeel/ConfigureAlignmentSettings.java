package com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.TextDirectionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConfigureAlignmentSettings {

    private static final String TAG = ConfigureAlignmentSettings.class.getName();

    /**
     * To align text horizontally, use the Style object's setHorizontalAlignment method.
     */
    public void horizontalTextAlignment() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            Style style = cell.getStyle();

            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            //Setting the horizontal alignment of the text in the "A1" cell
            style.setHorizontalAlignment(TextAlignmentType.CENTER);

            //Saved style
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "HorizontalTextAlignment_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Horizontal Text Alignment", e);
        }
    }

    /**
     * To align text vertically, use the Style object's setVerticalAlignment method.
     */
    public void verticalTextAlignment() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            // Setting the vertical alignment of the text in a cell
            Style style = cell.getStyle();
            style.setVerticalAlignment(TextAlignmentType.CENTER);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "VerticalTextAlignment_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Vertical Text Alignment", e);
        }
    }

    /**
     * To set the indentation level of text in a cell, use the Style object's setIndentLevel method.
     */
    public void indentation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            Style style = cell.getStyle();
            style.setIndentLevel(2);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "Indentation_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Indentation", e);
        }
    }

    /**
     * Set the text's orientation (rotation) in the cell with the Style object's setRotationAngle.
     */
    public void orientation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            //Setting the rotation of the text (inside the cell) to 25
            Style style = cell.getStyle();
            style.setRotationAngle(25);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "Orientation_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Orientation", e);
        }
    }

    /**
     * Set the text to be wrapped within the cell by using the setTextWrapped method of the Style object.
     * The Style object's setTextWrapped method to control text wrapping within a cell.
     */
    public void wrapText() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            //Enabling the text to be wrapped within the cell
            Style style = cell.getStyle();
            style.setTextWrapped(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "WrapText_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Wrap Text", e);
        }
    }

    /**
     * Text wrapping fits text into a cell by wrapping it vertically.
     * Another option for text that is too large for the cell it is in is to shrink it to fit.
     * This feature adjusts the size of the text until it fits in the cell. IT is particularly useful for relatively short labels - long strings might become difficult to read when shrunk.
     */
    public void shrinkToFit() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            //Shrinking the text to fit according to the dimensions of the cell
            Style style = cell.getStyle();
            style.setShrinkToFit(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "ShrinkToFit_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Shrink to Fit", e);
        }
    }

    /**
     * To merge cells into a single cell, use the Cell collection's merge method.
     */
    public void mergeCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Merging the first three columns in the first row to create a single cell
            cells.merge(0, 0, 1, 3);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "MergeCells_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Merge Cells", e);
        }
    }

    /**
     * Set the reading order (the visual order in which the characters, words etc. are displayed.
     * For example, English is a left to right language while Arabic is a right to left language)
     * of text inside cells with the Style object's setTextDirection method.
     */
    public void textDirection() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding the current system date to "A1" cell
            Cell cell = cells.get("A1");
            //Adding some value to the "A1" cell
            cell.setValue("Visit Aspose!");

            //Setting the text direction from right to left
            Style style = cell.getStyle();
            style.setTextDirection(TextDirectionType.RIGHT_TO_LEFT);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "TextDirection_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Text Direction", e);
        }
    }
}
