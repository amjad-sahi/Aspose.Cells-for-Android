package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class RenderCustomDateFormatPattern {

    private static final String TAG = RenderCustomDateFormatPattern.class.getName();

    public void renderCustomDateFormatPattern() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "DateFormat.xlsx");
            workbook.save(filePath + "RenderCustomDateFormatPattern_Out.pdf");
        } catch (Exception e) {
            Log.e(TAG, "Render Custom Date Format Pattern", e);
        }
    }
}
