package com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FontSettings {

    private static final String TAG = FontSettings.class.getName();

    /**
     * Apply any font to the text inside a cell with the Font object's Name method.
     */
    public void setFontName() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting the font name to "Times New Roman"
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setName("Times New Roman");

            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetFontName_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Font Name", e);
        }
    }

    /**
     * Set the font style to bold by providing true as a parameter of the Font object's setBold method.
     */
    public void setFontStyleToBold() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting the font weight to bold
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setBold(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetFontStyleToBold_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Font Style to Bold", e);
        }
    }

    /**
     * Set the font size using the Font object's setSize method.
     */
    public void setFontSize() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Set the font Size
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setSize(14);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetFontSize_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Font Size", e);
        }
    }

    /**
     * If you want the text to be underlines, use the Font object's Underline method.
     * Aspose.Cells offers various pre-defined font underline types in the form of the FontUnderlineType enumeration.
     */
    public void setFontUnderline() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting the font to be underlined
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setUnderline(FontUnderlineType.SINGLE);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetFontUnderline_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Font Underline Type", e);
        }
    }

    /**
     * Set the color of the font by using the Font object's setColor method.
     * Select any color from the Color enumeration and assign it color to the Font object's setColor method.
     */
    public void setFontColor() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting the font color to blue
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setColor(Color.getBlue());
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetFontColor_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Font Color", e);
        }
    }

    /**
     * Apply the strike out effect on the font by using the Font object's setStrikeout method.
     */
    public void setStrikeOutEffectOnFont() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting the strike out effect on the font
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setStrikeout(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetStrikeOutEffectOnFont_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set Strike Out Effect on Font", e);
        }
    }

    /**
     * Turn text into subscript with the Font object's setSubscript method.
     */
    public void setSubScriptEffectOnFont() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting subscript effect
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setSubscript(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetSubScriptEffectOnFont_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set SubScript Effect on Font", e);
        }
    }

    /**
     * Turn text into superscript using the Font object's setSuperscript method.
     */
    public void setSuperScriptEffectOnFont() {
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

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");
            cell.setValue("Hello Aspose!");

            //Setting superscript effect
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setSuperscript(true);
            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + File.separator + "SetSuperScriptEffectOnFont_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Set SuperScript Effect on Font", e);
        }
    }



}
