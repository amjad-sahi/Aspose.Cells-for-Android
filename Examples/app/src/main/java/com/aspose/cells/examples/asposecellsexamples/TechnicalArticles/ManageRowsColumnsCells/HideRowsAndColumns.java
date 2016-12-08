package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class HideRowsAndColumns {
    private static final String TAG = HideRowsAndColumns.class.getName();

    public void hideRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook();

            Worksheet worksheet = wb.getWorksheets().get(0);

            //Hide the 3rd row of the worksheet
            worksheet.getCells().hideRow(2);

            //Hide the 2nd column of the worksheet
            worksheet.getCells().hideColumn(1);

            wb.save(filePath + "Cells_HideRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Hiding Rows And Columns", e);
        }
    }
}
