package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DisplayRowsAndColumns {

    private static final String TAG = DisplayRowsAndColumns.class.getName();

    public void displayRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook();

            Worksheet worksheet = wb.getWorksheets().get(0);

            //Display the 3rd row of the worksheet and set its height 25
            worksheet.getCells().unhideRow(2, 25);

            //Display the 2nd column of the worksheet and set its width 15
            worksheet.getCells().unhideColumn(1, 15);

            wb.save(filePath + "Cells_DisplayRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Displaying Rows And Columns", e);
        }
    }
}
