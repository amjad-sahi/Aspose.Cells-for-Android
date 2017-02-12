package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class Managing3DEffectsforShapes {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Managing 3D Effects for Shapes
     */
    public static void Run(Context context) {
        Log.e(TAG, "Running Managing3DEffectsforShapes");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleManaging3DEffectsforShapes.xlsx");

            //Load a sample spreadsheet containing a shape
            Workbook book = new Workbook(in);

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Access first shape from the collection
            Shape shape = sheet.getShapes().get(0);

            //Get the instance of ThreeDFormat from the Shape object
            ThreeDFormat threeD = shape.getThreeDFormat();

            //Set its ContourWidth & ExtrusionHeight properties
            threeD.setContourWidth(15);
            threeD.setExtrusionHeight(30);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputManaging3DEffectsforShapes.xlsx");

            Log.e(TAG, "outputManaging3DEffectsforShapes.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occured in Managing3DEffectsforShapes");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
