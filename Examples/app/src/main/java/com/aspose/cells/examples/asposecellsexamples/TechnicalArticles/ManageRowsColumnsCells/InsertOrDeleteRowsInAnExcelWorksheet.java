package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class InsertOrDeleteRowsInAnExcelWorksheet {

    private static final String TAG = InsertOrDeleteRowsInAnExcelWorksheet.class.getName();

    public void insertAndDeleteRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a Workbook object.
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the first worksheet in the book.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Insert 10 rows at row index 2 (insertion starts at 3rd row)
            sheet.getCells().insertRows(2, 10);

            //Delete 5 rows now. (8th row - 12th row)
            sheet.getCells().deleteRows(7, 5, true);

            //Save the Excel file.
            workbook.save(filePath + "InsertAndDeleteRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Insert or Delete Rows in an Excel Worksheet", e);
        }
    }
}
