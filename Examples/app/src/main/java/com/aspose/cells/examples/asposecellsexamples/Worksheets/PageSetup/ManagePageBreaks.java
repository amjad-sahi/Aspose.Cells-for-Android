package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HorizontalPageBreakCollection;
import com.aspose.cells.VerticalPageBreakCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class ManagePageBreaks {

    private static final String TAG = ManagePageBreaks.class.getName();

    public void addPageBreaks() {
        // Instantiate a Workbook object
        Workbook workbook = new Workbook();

        // Add a page break at cell Y30
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet = worksheets.get(0);
        HorizontalPageBreakCollection hPageBreaks= worksheet.getHorizontalPageBreaks();
        hPageBreaks.add("Y30");
        VerticalPageBreakCollection vPageBreaks= worksheet.getVerticalPageBreaks();
        vPageBreaks.add("Y30");
    }

    public void clearAllPageBreaks() {
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();
        workbook.getWorksheets().get(0).getHorizontalPageBreaks().clear();
        workbook.getWorksheets().get(0).getVerticalPageBreaks().clear();
    }

    public void removeSpecificPageBreak() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;


            Workbook workbook = new Workbook(filePath + "PageBreaks.xls");
            Worksheet worksheet = workbook.getWorksheets().get(0);

            HorizontalPageBreakCollection hPageBreaks = worksheet.getHorizontalPageBreaks();
            hPageBreaks.removeAt(0);

            VerticalPageBreakCollection vPageBreaks = worksheet.getVerticalPageBreaks();
            vPageBreaks.removeAt(0);
        } catch (Exception e) {
            Log.e(TAG, "Remove Specific Page Break", e);
        }
    }

}
