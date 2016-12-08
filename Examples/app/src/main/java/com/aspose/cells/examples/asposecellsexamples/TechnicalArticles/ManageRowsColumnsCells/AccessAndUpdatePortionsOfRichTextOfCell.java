package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.FontSetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AccessAndUpdatePortionsOfRichTextOfCell {

    private static final String TAG = AccessAndUpdatePortionsOfRichTextOfCell.class.getName();

    public void accessAndUpdatePortionsOfRichTextOfCell() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "RichText.xlsx");

            Worksheet worksheet = workbook.getWorksheets().get(0);

            Cell cell = worksheet.getCells().get("A1");

            Log.i(TAG, "Before updating the font settings....");

            FontSetting[] fnts = cell.getCharacters();

            for (int i = 0; i < fnts.length; i++) {
                Log.i(TAG, fnts[i].getFont().getName());
            }

            //Modify the first FontSetting Font Name
            fnts[0].getFont().setName("Arial");

            //And update it using SetCharacters() method
            cell.setCharacters(fnts);

            Log.i(TAG, "After updating the font settings....");

            fnts = cell.getCharacters();

            for (int i = 0; i < fnts.length; i++) {
                Log.i(TAG, fnts[i].getFont().getName());
            }

            //Save workbook
            workbook.save("PortionsOfRichTextOfCell_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Access and Update the Portions of Rich Text of Cell", e);
        }
    }
}
