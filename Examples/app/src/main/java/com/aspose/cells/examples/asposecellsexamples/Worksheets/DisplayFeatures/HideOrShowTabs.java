package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class HideOrShowTabs {

    private static final String TAG = HideOrShowTabs.class.getName();

    public void hideTabs() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Hide the tabs of the Excel file
            workbook.getSettings().setShowTabs(false);

            workbook.save(filePath + File.separator + "HideTabs_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Hide Tabs", e);
        }
    }

    public void makeTabsVisible() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Display the tabs of the Excel file
            workbook.getSettings().setShowTabs(true);

            workbook.save(filePath + File.separator + "MakeTabsVisible_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Make Tabs Visible", e);
        }
    }


}
