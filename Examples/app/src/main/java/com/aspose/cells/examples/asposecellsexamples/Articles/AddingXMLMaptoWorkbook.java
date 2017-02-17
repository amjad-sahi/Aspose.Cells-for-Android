package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class AddingXMLMaptoWorkbook {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Adding XML Map to Workbook
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running AddingXMLMaptoWorkbook");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            // Create a workbook
            Workbook book = new Workbook();

            // URL that contains your XML data for mapping
            String XML = "https://docs.aspose.com/download/attachments/5018589/sampleXML.txt";

            // Import your XML Map data starting from cell A1
            book.importXml(XML, "Sheet1", 0, 0);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputAddingXMLMaptoWorkbook.xlsx");

            Log.w(TAG, "outputAddingWordArtwithBuiltinStyles.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in AddingXMLMaptoWorkbook");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }
    }
}
