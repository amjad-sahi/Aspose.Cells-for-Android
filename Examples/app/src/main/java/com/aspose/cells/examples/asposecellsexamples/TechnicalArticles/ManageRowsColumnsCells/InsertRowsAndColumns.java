package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class InsertRowsAndColumns {

    private static final String TAG = InsertRowsAndColumns.class.getName();

    public void insertRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook wb = new Workbook();

            Worksheet worksheet = wb.getWorksheets().get(0);

            Cells cells = worksheet.getCells();

            //Put some values into cells
            Cell cell = cells.get("A1");
            cell.putValue("Aspose");
            cell = cells.get("A2");
            cell.putValue(123);
            cell = cells.get("A3");
            cell.putValue("Hello World");
            cell = cells.get("B1");
            cell.putValue(120);

            //Insert a row or column into the worksheet

            //Insert 10 rows starting from 3rd row
            cells.insertRows(2, 10);

            //Insert 1 column starting from 2nd column
            cells.insertColumns(1, 1);

            wb.save(filePath + "Cells_InsertRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Insert Rows And Columns", e);
        }
    }
}
