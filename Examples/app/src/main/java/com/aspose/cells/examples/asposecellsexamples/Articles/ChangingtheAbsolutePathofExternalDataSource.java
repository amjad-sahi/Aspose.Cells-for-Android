package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ChangingtheAbsolutePathofExternalDataSource {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Changing the Absolute Path of External Data Source
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running ChangingtheAbsolutePathofExternalDataSource");

        try {
            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleChangingtheAbsolutePathofExternalDataSource.xlsx");

            //Load your source excel file containing the external link
            Workbook wb = new Workbook(in);

            //Access the first external link
            ExternalLink externalLink = wb.getWorksheets().getExternalLinks().get(0);

            //Print the data source of external link, it will print existing remote path
            System.out.println("External Link Data Source: " + externalLink.getDataSource());

            // Remove the remote path and print the new data source
            // Assign the new data source to external link and print again, it will now print data source with local path
            externalLink.setDataSource("ExternalAccounts.xlsx");
            Log.w(TAG, "External Link Data Source After Removing Remote Path: " + externalLink.getDataSource());

            // Change the absolute path of the workbook, it will also change the external link path
            wb.setAbsolutePath("C:\\Files\\Extra\\");

            // Now print the data source again
            Log.w(TAG, "External Link Data Source After Changing Workbook.AbsolutePath to Local Path: " + externalLink.getDataSource());

            // Change the absolute path of the workbook to some remote path, it will again affect the external link path
            wb.setAbsolutePath("http://www.aspose.com/WebFiles/ExcelFiles/");

            // Now print the data source again
            Log.w(TAG, "External Link Data Source After Changing Workbook.AbsolutePath to Remote Path: " + externalLink.getDataSource());
            Log.w(TAG, "------------------------------------------");

            Log.w(TAG, "ChangingtheAbsolutePathofExternalDataSource executed successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in ChangingtheAbsolutePathofExternalDataSource");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
