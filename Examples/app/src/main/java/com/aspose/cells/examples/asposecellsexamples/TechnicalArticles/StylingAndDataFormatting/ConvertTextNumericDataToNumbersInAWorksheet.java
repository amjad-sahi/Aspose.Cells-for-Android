package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class ConvertTextNumericDataToNumbersInAWorksheet {
    private static final String TAG = ConvertTextNumericDataToNumbersInAWorksheet.class.getName();

    public void convertTextNumericDataToNumbersInAWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load an existing spreadsheet in an instance of Workbook
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Loop over the collection of Worksheets
            for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
                //Convert all numeric string data to numbers
                workbook.getWorksheets().get(i).getCells().convertStringToNumericValue();
            }

            //Save the result
            workbook.save(filePath + "ConvertTextNumericData_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Convert Text Numeric Data to Numbers in a Worksheet", e);
        }
    }

}
