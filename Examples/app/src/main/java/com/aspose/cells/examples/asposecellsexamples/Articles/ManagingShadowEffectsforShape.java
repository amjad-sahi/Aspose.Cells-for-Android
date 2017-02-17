package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class ManagingShadowEffectsforShape {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Managing Shadow Effects for Shape
     */
    public static void Run(Context context)
    {
        Log.w(TAG, "Running ManagingShadowEffectsforShape");

        try
        {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleManagingShadowEffectsforShape.xlsx");

            //Load a sample spreadsheet containing a shape
            Workbook book = new Workbook(in);

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Access first shape from the collection
            Shape shape = sheet.getShapes().get(0);

            //Get the instance of ShadowEffect from the Shape object
            ShadowEffect shadow = shape.getShadowEffect();

            //Set its Angle, Blur, Size, Transparency and Distance properties
            shadow.setAngle(150);
            shadow.setBlur(30);
            shadow.setSize(0.9);
            shadow.setTransparency(0.5);
            shadow.setDistance(80);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputManagingShadowEffectsforShape.xlsx");

            Log.w(TAG, "outputManagingShadowEffectsforShape.xlsx created successfully");
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Some exception occurred in ManagingShadowEffectsforShape");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
