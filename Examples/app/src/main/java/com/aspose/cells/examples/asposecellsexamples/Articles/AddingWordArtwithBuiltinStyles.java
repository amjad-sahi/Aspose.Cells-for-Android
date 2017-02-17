package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class AddingWordArtwithBuiltinStyles {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Adding WordArt with Builtin Styles
     */
    public static void Run(Context context) {
        Log.w(TAG, "Running AddingWordArtwithBuiltinStyles");

        try {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Access ShapeCollection of first worksheet
            ShapeCollection shapes = sheet.getShapes();

            //Add WordArt with built-in styles
            shapes.addWordArt(PresetWordArtStyle.WORD_ART_STYLE_1, "Aspose File Format APIs", 00, 0, 0, 0, 100, 800);
            shapes.addWordArt(PresetWordArtStyle.WORD_ART_STYLE_2, "Aspose File Format APIs", 10, 0, 0, 0, 100, 800);
            shapes.addWordArt(PresetWordArtStyle.WORD_ART_STYLE_3, "Aspose File Format APIs", 20, 0, 0, 0, 100, 800);
            shapes.addWordArt(PresetWordArtStyle.WORD_ART_STYLE_4, "Aspose File Format APIs", 30, 0, 0, 0, 100, 800);
            shapes.addWordArt(PresetWordArtStyle.WORD_ART_STYLE_5, "Aspose File Format APIs", 40, 0, 0, 0, 100, 800);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputAddingWordArtwithBuiltinStyles.xlsx");

            Log.w(TAG, "outputAddingWordArtwithBuiltinStyles.xlsx created successfully");
        } catch (Exception ex) {
            Log.e(TAG, "Some exception occurred in AddingWordArtwithBuiltinStyles");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }
    }
}
