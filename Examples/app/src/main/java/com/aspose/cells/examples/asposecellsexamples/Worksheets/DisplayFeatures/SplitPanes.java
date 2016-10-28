package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class SplitPanes {

    private static final String TAG = SplitPanes.class.getName();

    public void splitPanes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Set the active cell
            workbook.getWorksheets().get(0).setActiveCell("A20");
            //Split the worksheet window
            workbook.getWorksheets().get(0).split();

            workbook.save(filePath + File.separator + "SplitPanes_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Split Panes", e);
        }
    }

    public void removePanes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Set the active cell
            workbook.getWorksheets().get(0).setActiveCell("A20");

            //Remove split panes
            workbook.getWorksheets().get(0).removeSplit();

            workbook.save(filePath + File.separator + "RemovePanes_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Remove Panes", e);
        }
    }

}
