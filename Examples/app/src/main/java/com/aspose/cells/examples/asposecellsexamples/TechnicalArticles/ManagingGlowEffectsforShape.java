package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ManagingGlowEffectsforShape {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Managing Glow Effects for Shapes
     */
    public static void Run(Context context) {
        Log.e(TAG, "Running ManagingGlowEffectsforShape");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleManagingGlowEffectsforShape.xlsx");

            //Load a sample spreadsheet containing a shape
            Workbook book = new Workbook(in);

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Access first shape from the collection
            Shape shape = sheet.getShapes().get(0);

            //Get the instance of GlowEffect from the Shape object
            GlowEffect glow = shape.getGlow();

            //Set its Size & Transparency properties
            glow.setSize(90);
            glow.setTransparency(0.5);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputManagingGlowEffectsforShape.xlsx");

            Log.e(TAG, "outputManagingGlowEffectsforShape.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occured in ManagingGlowEffectsforShape");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
