package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ApplySuperscriptAndSubscriptEffectsOnFonts {
    private static final String TAG = ApplySuperscriptAndSubscriptEffectsOnFonts.class.getName();

    public void setSuperscriptEffect() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            Cells cells = worksheet.getCells();

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");

            cell.setValue("Hello");

            //Setting the font name to "Times New Roman"
            Style style = cell.getStyle();

            Font font = style.getFont();
            font.setSuperscript(true);

            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + "Superscript_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Superscript Effect", e);
        }
    }

    public void setSubscriptEffect() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Accessing the added worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            Cells cells = worksheet.getCells();

            //Adding some value to the "A1" cell
            Cell cell = cells.get("A1");

            cell.setValue("Hello");

            //Setting the font name to "Times New Roman"
            Style style = cell.getStyle();

            Font font = style.getFont();
            font.setSubscript(true);

            cell.setStyle(style);

            //Saving the modified Excel file in default format
            workbook.save(filePath + "Subscript_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Subscript Effect", e);
        }
    }
}
