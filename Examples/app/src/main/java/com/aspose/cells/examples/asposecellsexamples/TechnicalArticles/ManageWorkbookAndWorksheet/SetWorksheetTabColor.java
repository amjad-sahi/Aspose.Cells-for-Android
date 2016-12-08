package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Color;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetWorksheetTabColor {

    private static final String TAG = SetWorksheetTabColor.class.getName();

    public void worksheetTabColor() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the first worksheet in the book
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Set the tab color
            worksheet.setTabColor(Color.getRed());

            //Save the Excel file
            workbook.save(filePath + "WorksheetTabColor_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Worksheet Tab Color", e);
        }
    }
}
