package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DeleteRowsAndColumns {

    private static final String TAG = DeleteRowsAndColumns.class.getName();

    public void deleteRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook();

            Worksheet worksheet = wb.getWorksheets().get(0);

            Cells cells = worksheet.getCells();

            //Put some values into cells
            Cell cell = cells.get("A1");
            cell.putValue("Row-1");

            cell = cells.get("A2");
            cell.putValue("Row-2");

            cell = cells.get("A3");
            cell.putValue("Row-3");

            cell = cells.get("A4");
            cell.putValue("Row-4");

            cell = cells.get("A5");
            cell.putValue("Row-5");

            cell = cells.get("B1");
            cell.putValue("Column B");

            cell = cells.get("C1");
            cell.putValue("Column C");

            cell = cells.get("D1");
            cell.putValue("Column D");

            //Delete 2 rows starting from 3rd row i.e 3rd and 4th rows
            cells.deleteRows(2, 2, false);
            //Delete 1 column starting from 2nd column i.e column B
            cells.deleteColumns(1, 1, false);
            wb.save(filePath + "Cells_DeleteRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Delete Rows And Columns", e);
        }
    }
}
