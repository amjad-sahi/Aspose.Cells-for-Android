package com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ColorAndPalette {

    private static final String TAG = ColorAndPalette.class.getName();

    /**
     * Aspose.Cells also supports a palette of 56 colors.
     * If you want to use a custom color that is not defined in the palette, add that color to the palette before using it.
     * This example demonstrates how to add a custom color to the palette before applying it on a font.
     */
    public void addCustomColorsToPalette() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding custom color to the palette at 55th index
            Color color = Color.fromArgb(212,213,0);
            workbook.changePalette(color,55);

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Accessing the "A1" cell from the worksheet
            Cell cell = worksheet.getCells().get("A1");

            //Adding some value to the "A1" cell
            cell.setValue("Hello Aspose!");

            //Setting the custom color to the font
            Style style = cell.getStyle();
            Font font = style.getFont();
            font.setColor(color);

            cell.setStyle(style);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddCustomColorsToPalette_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Add Custom Colors to Palette", e);
        }
    }
}
