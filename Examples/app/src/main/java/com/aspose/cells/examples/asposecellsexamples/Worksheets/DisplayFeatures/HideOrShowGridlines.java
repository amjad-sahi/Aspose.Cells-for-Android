package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class HideOrShowGridlines {

    private static final String TAG = HideOrShowGridlines.class.getName();

    public void makeGridlinesVisible() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Display the gridlines of the worksheet
            worksheet.setGridlinesVisible(true);

            workbook.save(filePath + File.separator + "MakeGridlinesVisible_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Make Gridlines Visible", e);
        }
    }

    public void hideGridlines() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Hide gridlines of the worksheet
            worksheet.setGridlinesVisible(false);

            workbook.save(filePath + File.separator + "HideGridlines_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Hide Gridlines", e);
        }
    }



}
