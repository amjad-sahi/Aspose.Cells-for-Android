package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class HideOrShowScrollBars {
    private static final String TAG = HideOrShowScrollBars.class.getName();

    public void hideScrollBars() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Hide the vertical scroll bar of the Excel file
            workbook.getSettings().setVScrollBarVisible(false);

            //Hide the horizontal scroll bar of the Excel file
            workbook.getSettings().setHScrollBarVisible(false);

            workbook.save(filePath + File.separator + "HideScrollBars_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Hide Scroll Bars", e);
        }
    }

    public void makeScrollBarsVisible() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Display the vertical scroll bar of the Excel file
            workbook.getSettings().setVScrollBarVisible(true);

            //Display the horizontal scroll bar of the Excel file
            workbook.getSettings().setHScrollBarVisible(true);

            workbook.save(filePath + File.separator + "MakeScrollBarsVisible_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Make ScrollBarsVisible Scroll Bars", e);
        }
    }


}
