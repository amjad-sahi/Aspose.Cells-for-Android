package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ManagingReflectionEffectsforShapes {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Managing Reflection Effects for Shapes
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running ManagingReflectionEffectsforShapes");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleManagingReflectionEffectsforShapes.xlsx");

            //Load a sample spreadsheet containing a shape
            Workbook book = new Workbook(in);

            // Access first worksheet
            Worksheet ws = book.getWorksheets().get(0);

            // Access first shape
            Shape sh = ws.getShapes().get(0);

            // Set the reflection effect of the shape
            // Set its Blur, Size, Transparency and Distance properties
            ReflectionEffect re = sh.getReflection();
            re.setBlur(30);
            re.setSize(90);
            re.setTransparency(0);
            re.setDistance(80);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputManagingReflectionEffectsforShapes.xlsx");

            Log.w(TAG, "outputManagingReflectionEffectsforShapes.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in AddingWordArtwithBuiltinStyles");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }
    }
}
