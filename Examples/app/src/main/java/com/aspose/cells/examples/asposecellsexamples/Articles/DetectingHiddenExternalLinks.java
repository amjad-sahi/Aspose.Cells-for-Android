package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class DetectingHiddenExternalLinks {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Detecting Hidden External Links
     */
    public static void Run(Context context)
    {
        Log.w(TAG, "Running DetectingHiddenExternalLinks");

        try
        {
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleDetectingHiddenExternalLinks.xlsx");

            //Loads the workbook which contains hidden external links
            Workbook book = new Workbook(in);

            //Access the external link collection of the workbook
            ExternalLinkCollection links = book.getWorksheets().getExternalLinks();

            //Print all the external links and check the IsVisible property
            for (int i = 0; i < links.getCount(); i++)
            {
                Log.w(TAG, "Data Source: " + links.get(i).getDataSource());
                Log.w(TAG, "Is Referred: " + links.get(i).isReferred());
                Log.w(TAG, "Is Visible: " + links.get(i).isVisible());
                Log.w(TAG, "-------------------------------");
            }

            Log.w(TAG, "DetectingHiddenExternalLinks executed successfully");
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Some exception occurred in DetectingHiddenExternalLinks");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
