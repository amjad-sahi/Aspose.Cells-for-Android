package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class HideAndShowRowsAndColumns {

    private static final String TAG = HideAndShowRowsAndColumns.class.getName();

    public void hideRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Hiding the 3rd row of the worksheet
            cells.hideRow(2);

            //Hiding the 2nd column of the worksheet
            cells.hideColumn(1);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "HideRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Hide Rows And Columns", e);
        }
    }

    public void showRowsAndColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            //Unhiding the 3rd row and setting its height to 13.5
            cells.unhideRow(2,13.5);

            //Unhiding the 2nd column and setting its width to 8.5
            cells.unhideColumn(1,8.5);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "ShowRowsAndColumns_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Show Rows And Columns", e);
        }
    }


}
