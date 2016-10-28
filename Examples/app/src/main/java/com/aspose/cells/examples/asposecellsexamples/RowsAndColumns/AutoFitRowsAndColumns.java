package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AutoFitRowsAndColumns {

    private static final String TAG = AutoFitRowsAndColumns.class.getName();

    public void autoFitRow() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto-fitting the 3rd row of the worksheet
            worksheet.autoFitRow(2);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "AutoFitRow_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Auto-fit Row", e);
        }
    }

    public void autoFitRowInARangeOfCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto-fitting the 3rd row of the worksheet based on the contents in a range of
            //cells (from 1st to 9th column) within the row
            worksheet.autoFitRow(2,0,8);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "AutoFitRowInARangeOfCells_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Auto-fit Row in a range of cells", e);
        }
    }

    public void autoFitColumn() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto-fitting the 4th column of the worksheet
            worksheet.autoFitColumn(3);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "AutoFitColumn_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Auto-fit Column", e);
        }
    }

    public void autoFitColumnInARangeOfCells() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto-fitting the 4th column of the worksheet based on the contents in a range of
            //cells (from 1st to 9th row) within the column
            worksheet.autoFitColumn(3,0,8);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "AutoFitColumnInARangeOfCells_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Auto-fit Column in a range of cells", e);
        }
    }
}
