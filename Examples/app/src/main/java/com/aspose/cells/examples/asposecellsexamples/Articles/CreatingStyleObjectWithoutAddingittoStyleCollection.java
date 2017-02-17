package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class CreatingStyleObjectWithoutAddingittoStyleCollection {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Creating Style Object Without Adding it to Style Collection
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running CreatingStyleObjectWithoutAddingittoStyleCollection");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            // Create a Style object using CellsFactory class
            CellsFactory cf = new CellsFactory();
            Style st = cf.createStyle();

            // Set the Style fill color to Yellow
            st.setPattern(BackgroundType.SOLID);
            st.setForegroundColor(Color.getYellow());

            // Create a workbook and set its default style using the created Style object
            Workbook wb = new Workbook();
            wb.setDefaultStyle(st);

            //Save the result in XLSX format
            wb.save(SD_PATH + "outputCreatingStyleObjectWithoutAddingittoStyleCollection.xlsx");

            Log.w(TAG, "outputCreatingStyleObjectWithoutAddingittoStyleCollection.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in CreatingStyleObjectWithoutAddingittoStyleCollection");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
