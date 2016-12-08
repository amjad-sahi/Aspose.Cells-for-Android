package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class UngroupRowsAndColumns {

    private static final String TAG = UngroupRowsAndColumns.class.getName();

    public void ungroupRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook(filePath + "Book1.xls");
            Worksheet worksheet = wb.getWorksheets().get(0);

            Cells cells = worksheet.getCells();
            cells.ungroupRows(0, 9);
            cells.ungroupColumns(0, 1);

            wb.save(filePath + "UngroupRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Ungrouping Rows And Columns", e);
        }
    }

}
