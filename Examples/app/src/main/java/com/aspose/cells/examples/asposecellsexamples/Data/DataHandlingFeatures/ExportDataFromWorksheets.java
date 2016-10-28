package com.aspose.cells.examples.asposecellsexamples.Data.DataHandlingFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;
import java.io.FileInputStream;

public class ExportDataFromWorksheets {

    private static final String TAG = ExportDataFromWorksheets.class.getName();

    public void exportDataToArray() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating a file stream containing the Excel file to be opened
            FileInputStream fstream = new FileInputStream(filePath + File.separator + "Book1.xls");

            //Instantiating a Workbook object
            Workbook workbook = new Workbook(fstream);

            //Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Exporting the contents of 7 rows and 2 columns starting from 1st cell to Array.
            Object dataTable [][] = worksheet.getCells().exportArray(0, 0, 7, 2);

            //Closing the file stream to free all resources
            fstream.close();
        } catch (Exception e) {
            Log.e(TAG, "Export Data to Array", e);
        }
    }

}
