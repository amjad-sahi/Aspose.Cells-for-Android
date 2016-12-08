package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GroupRowsAndColumns {

    private static final String TAG = GroupRowsAndColumns.class.getName();

    public void groupRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook(filePath + "Book1.xls");
            Worksheet worksheet = wb.getWorksheets().get(0);

            Cells cells = worksheet.getCells();
            cells.groupRows(0, 9, false);
            cells.groupColumns(0, 1, false);

            //Set SummaryRowBelow property
            worksheet.getOutline().SummaryRowBelow = true;

            //Set SummaryColumnRight property
            worksheet.getOutline().SummaryColumnRight = true;

            wb.save(filePath + "GroupRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Grouping Rows And Columns", e);
        }
    }
}
