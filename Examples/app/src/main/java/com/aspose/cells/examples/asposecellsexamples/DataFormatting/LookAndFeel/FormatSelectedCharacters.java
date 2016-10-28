package com.aspose.cells.examples.asposecellsexamples.DataFormatting.LookAndFeel;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class FormatSelectedCharacters {

    private static final String TAG = FormatSelectedCharacters.class.getName();

    /**
     * Formatting Selected Characters.
     */
    public void formatSelectedCharacters() {
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
            cell.setValue("Visit Aspose!");

            Font font = cell.characters(6, 7).getFont();

            //Setting the font of selected characters to bold
            font.setBold(true);

            //Setting the font color of selected characters to blue
            font.setColor(Color.getBlue());

            //Saving the Excel file
            workbook.save(filePath + File.separator + "FormatSelectedCharacters_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Format Selected Characters", e);
        }
    }
}
