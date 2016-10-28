package com.aspose.cells.examples.asposecellsexamples.Table;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class ConvertTableToRangeOfData {

    private static final String TAG = ConvertTableToRangeOfData.class.getName();

    public void convertTableToRange() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Open an existing file that contains a table/list object in it
            Workbook wb = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Convert the first table/list object (from the first worksheet) to normal range
            wb.getWorksheets().get(0).getListObjects().get(0).convertToRange();

            //Save the file
            wb.save(filePath + File.separator + "ConvertTableToRange_Out.xlsx");

        } catch (Exception e) {
            Log.e(TAG, "Convert a Table to a Range", e);
        }
    }
}
