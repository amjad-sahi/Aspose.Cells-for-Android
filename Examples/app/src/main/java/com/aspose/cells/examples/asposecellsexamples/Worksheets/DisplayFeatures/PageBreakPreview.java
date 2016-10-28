package com.aspose.cells.examples.asposecellsexamples.Worksheets.DisplayFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class PageBreakPreview {
    private static final String TAG = PageBreakPreview.class.getName();

    public void enableNormalView() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Displaying the worksheet in page break preview
            worksheet.setPageBreakPreview(false);

            workbook.save(filePath + File.separator + "EnableNormalView_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Enable Normal View", e);
        }
    }

    public void enablePageBreakPreview() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Adding a new worksheet to the Workbook object
            WorksheetCollection worksheets = workbook.getWorksheets();

            Worksheet worksheet = worksheets.get(0);

            //Displaying the worksheet in page break preview
            worksheet.setPageBreakPreview(true);

            workbook.save(filePath + File.separator + "EnablePageBreakPreview_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Enable Page Break Preview", e);
        }
    }

}
