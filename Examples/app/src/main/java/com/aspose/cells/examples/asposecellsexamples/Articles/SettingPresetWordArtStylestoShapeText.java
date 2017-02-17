package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class SettingPresetWordArtStylestoShapeText {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Setting Preset WordArt Styles to Shape Text
     */
    public static void Run(Context context)
    {
        Log.w(TAG, "Running SettingPresetWordArtStylestoShapeText");

        try
        {
            //Get the path of Aspose directory inside the SD Card
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Create workbook object
            Workbook book = new Workbook();

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Create a TextBox with some text
            int index = sheet.getTextBoxes().add(0, 0, 100, 700);
            TextBox textBox = (TextBox)sheet.getShapes().get(index);
            textBox.setText("Aspose File Format APIs");
            textBox.getFont().setSize(44);

            //Set preset WordArt style to the text of the shape
            FontSetting fntSetting = (FontSetting)textBox.getCharacters().get(0);
            fntSetting.setWordArtStyle(PresetWordArtStyle.WORD_ART_STYLE_15);

            //Save the result in XLSX format
            book.save(SD_PATH + "outputSettingPresetWordArtStylestoShapeText.xlsx");

            Log.w(TAG, "outputSettingPresetWordArtStylestoShapeText.xlsx created successfully");
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Some exception occurred in SettingPresetWordArtStylestoShapeText");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
