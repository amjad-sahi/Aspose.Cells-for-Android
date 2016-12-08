package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.util.Iterator;

public class DetectEmptyWorksheets {

    private static final String TAG = DetectEmptyWorksheets.class.getName();

    public void detectEmptyWorksheets() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook book = new Workbook(filePath + "sample.xlsx");
            //Loop over all worksheets in the workbook
            for (int i = 0; i < book.getWorksheets().getCount(); i++) {
                Worksheet sheet = book.getWorksheets().get(i);
                // Check if worksheet has populated cells
                if (sheet.getCells().getMaxDataRow() != -1) {
                    Log.i(TAG, sheet.getName() + " is not empty because one or more cells are populated");
                }
                // Check if worksheet has shapes
                else if (sheet.getShapes().getCount() > 0) {
                    Log.i(TAG, sheet.getName() + " is not empty because there are one or more shapes");
                }
                // Check if worksheet has empty initialized cells
                else {
                    Range range = sheet.getCells().getMaxDisplayRange();
                    Iterator rangeIterator = range.iterator();
                    if (rangeIterator.hasNext()) {
                        Log.i(TAG, sheet.getName() + " is not empty because one or more cells are initialized");
                    }
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Detecting Empty Worksheets", e);
        }
    }

}
