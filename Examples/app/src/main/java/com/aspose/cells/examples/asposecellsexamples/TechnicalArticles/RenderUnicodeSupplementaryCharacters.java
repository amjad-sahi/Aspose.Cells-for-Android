package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class RenderUnicodeSupplementaryCharacters {

    private static final String TAG = RenderUnicodeSupplementaryCharacters.class.getName();

    public void renderUnicodeSupplementaryCharactersInOutputPdf() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load your source excel file containing Unicode Supplementary characters
            Workbook wb = new Workbook(filePath + "unicode-supplementary-characters.xlsx");

            //Save the workbook
            wb.save(filePath + "RenderUnicodeCharacters_Out.pdf");
        } catch (Exception e) {
            Log.e(TAG, "Render Unicode Supplementary characters in output Pdf", e);
        }
    }
}
