package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ExportingXMLMapData {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Exporting XML Map Data
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running ExportingXMLMapData");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleExportingXMLMapData.xlsx");

            //Load source workbook
            Workbook wb = new Workbook(in);

            //Export all XML data from all XML Maps inside the Workbook
            for (int i = 0; i < wb.getWorksheets().getXmlMaps().getCount(); i++)
            {
                //Access the XML Map
                XmlMap map = wb.getWorksheets().getXmlMaps().get(i);

                //Exports its XML Data
                wb.exportXml(map.getName(), SD_PATH + map.getName() + ".xml");
            }

            Log.w(TAG, "ExportingXMLMapData executed successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in ExportingXMLMapData");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
