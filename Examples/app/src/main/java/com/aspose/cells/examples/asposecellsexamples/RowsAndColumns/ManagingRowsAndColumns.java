package com.aspose.cells.examples.asposecellsexamples.RowsAndColumns;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ManagingRowsAndColumns {

    private static final String TAG = ManagingRowsAndColumns.class.getName();

    public void insertRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Inserting a row into the worksheet at 3rd position
            worksheet.getCells().insertRows(2,1);

            //Inserting 10 rows into the worksheet starting from 3rd row
            //worksheet.getCells().insertRows(2,10);

            //Saving the modified Excel file in default (that is Excel 2000) format
            workbook.save(filePath + File.separator + "InsertRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Insert Rows", e);
        }
    }

    public void deleteRows() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Deleting 3rd row from the worksheet
            worksheet.getCells().deleteRows(2,1,true);

            //Deleting 10 rows from the worksheet starting from 3rd row
            //worksheet.getCells().deleteRows(2,10,true);

            //Saving the modified Excel file in default (that is Excel 2000) format
            workbook.save(filePath + File.separator + "InsertRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Delete Rows", e);
        }
    }

    public void insertColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Inserting a column into the worksheet at 2nd position
            worksheet.getCells().insertColumns(1,1);

            //Saving the modified Excel file in default (that is Excel 2000) format
            workbook.save(filePath + File.separator + "InsertRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Insert Columns", e);
        }
    }

    public void deleteColumns() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Deleting a column from the worksheet at 2nd position
            worksheet.getCells().deleteColumns(1,1,true);

            //Saving the modified Excel file in default (that is Excel 2000) format
            workbook.save(filePath + File.separator + "InsertRows_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Delete Columns", e);
        }
    }

}
