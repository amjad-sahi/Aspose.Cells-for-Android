package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.SmartMarkers;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class ExportDataBarColorScaleAndIconSetConditionalFormatting {

    private static final String TAG = ExportDataBarColorScaleAndIconSetConditionalFormatting.class.getName();

    /**
     * You can export DataBar, ColorScale and IconSet Conditional Formatting while converting your Excel file into HTML.
     * This feature is partially supported by Microsoft Excel but Aspose.Cells supports it fully.
     */
    public void exportDataBarColorScaleAndIconSetConditionalFormatting() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load your sample excel file in a workbook object
            Workbook workbook = new Workbook(filePath + "ExportDataBar.xlsx");

            //Save it in HTML format
            workbook.save(filePath + "ExportDataBarColorScaleAndIconSet_Out.html", SaveFormat.HTML);

        } catch (Exception e) {
            Log.e(TAG, "Export DataBar, ColorScale and IconSet Conditional Formatting while Excel to HTML Conversion", e);
        }
    }
}