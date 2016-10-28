package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class HideOrShowRowColumnHeaders {
    private static final String TAG = HideOrShowRowColumnHeaders.class.getName();

    public void hideRowAndColumnHeaders() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            //Hide the headers of rows and columns
            worksheet.setRowColumnHeadersVisible(false);

            workbook.save(filePath + File.separator + "HideRowAndColumnHeaders_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Hide row and column headers", e);
        }
    }

    public void showRowAndColumnHeaders() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet worksheet = worksheets.get(0);

            //Display the headers of rows and columns
            worksheet.setRowColumnHeadersVisible(true);

            workbook.save(filePath + File.separator + "ShowRowAndColumnHeaders_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Hide row and column headers", e);
        }
    }
}
