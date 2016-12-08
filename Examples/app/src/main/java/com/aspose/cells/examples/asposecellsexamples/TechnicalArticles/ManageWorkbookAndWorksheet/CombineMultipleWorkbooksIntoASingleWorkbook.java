package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class CombineMultipleWorkbooksIntoASingleWorkbook {

    private static final String TAG = CombineMultipleWorkbooksIntoASingleWorkbook.class.getName();

    public void combineMultipleWorkbooksIntoASingleWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Open the first excel file.
            Workbook sourceBook1 = new Workbook(filePath + "PieBars.xlsx");

            //Define the second source book.
            //Open the second excel file.
            Workbook sourceBook2 = new Workbook(filePath + "Sample-oleobject.xlsx");

            //Combining the two workbooks
            sourceBook1.combine(sourceBook2);

            //Save the target book file.
            sourceBook1.save(filePath + "Combined.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Combine Multiple Workbooks into a Single Workbook", e);
        }
    }
}