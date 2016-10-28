package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.VisibilityType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class HideOrShowAWorksheet {
    private static final String TAG = HideOrShowAWorksheet.class.getName();

    public void hideAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Hide the worksheet of the Excel file
            worksheet.setVisible(false);

        } catch (Exception e) {
            Log.e(TAG, "Hide a Worksheet", e);
        }
    }

    public void makeAWorksheetVisible() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Display the worksheet of the Excel file
            worksheet.setVisible(true);

            //Saving the modified Excel file in default (that is Excel 2003) format
            workbook.save(filePath + File.separator + "MakeAWorksheetVisible_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Make a Worksheet Visible", e);
        }
    }

    public void setVisibilityType() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Hiding the worksheet of the Excel file
            worksheet.setVisibilityType(VisibilityType.VERY_HIDDEN);

        } catch (Exception e) {
            Log.e(TAG, "Set Visibility Type", e);
        }
    }


}
