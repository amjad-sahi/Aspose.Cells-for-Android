package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AccessAWorksheetMaximumDisplayRange {

    private static final String TAG = AccessAWorksheetMaximumDisplayRange.class.getName();

    public void accessAWorksheetMaximumDisplayRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a workbook from source file
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Access the first workbook
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the Maximum Display Range
            Range range = worksheet.getCells().getMaxDisplayRange();

            //Print the Maximum Display Range RefersTo property
            Log.v(TAG, "Maximum Display Range: " + range.getRefersTo());

        } catch (Exception e) {
            Log.e(TAG, "Access a Worksheet's Maximum Display Range", e);
        }
    }

}
