package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CopyRowsAndColumns {

    private static final String TAG = CopyRowsAndColumns.class.getName();

    public void copyRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new workbook
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the first worksheet cells
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Apply formulas to the cells
            for (int i = 0; i < 5; i++) {
                cells.get(0, i).setFormula("=Input!" + cells.get(0, i).getName());
            }

            //Copy the first row to next 10 rows
            for (int i = 1; i <= 10; i++) {
                cells.copyRow(cells, 0, i);
            }

            //Save the Excel file
            workbook.save(filePath + "CopyRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Copy Rows", e);
        }
    }

    public void copyColumns () {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the Cells collection
            Cells cells = worksheet.getCells();

            //Copy the first column to the third column
            cells.copyColumn(cells, 0, 2);

            //Save the Excel file
            workbook.save(filePath + "CopyColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Copy Columns", e);
        }
    }
}
