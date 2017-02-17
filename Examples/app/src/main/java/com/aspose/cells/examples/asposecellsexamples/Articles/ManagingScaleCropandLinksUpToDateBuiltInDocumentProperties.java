package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ManagingScaleCropandLinksUpToDateBuiltInDocumentProperties {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Managing ScaleCrop and LinksUpToDate Built In Document Properties
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running ManagingScaleCropandLinksUpToDateBuiltInDocumentProperties");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Create workbook
            Workbook wb = new Workbook();

            //Setting ScaleCrop and LinksUpToDate BuiltInDocumentProperties
            wb.getBuiltInDocumentProperties().setScaleCrop(true);
            wb.getBuiltInDocumentProperties().setLinksUpToDate(true);

            //Save the result in XLSX format
            wb.save(SD_PATH + "outputManagingScaleCropandLinksUpToDateBuiltInDocumentProperties.xlsx");

            Log.w(TAG, "outputManagingScaleCropandLinksUpToDateBuiltInDocumentProperties.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in ChangingtheAbsolutePathofExternalDataSource");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
