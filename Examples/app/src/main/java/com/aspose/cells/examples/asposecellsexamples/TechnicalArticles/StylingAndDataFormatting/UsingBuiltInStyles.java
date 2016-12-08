package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BuiltinStyleType;
import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;

import java.io.File;

public class UsingBuiltInStyles {

    private static final String TAG = UsingBuiltInStyles.class.getName();

    public void usingBuiltInStyles() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook();
            Style style = workbook.createBuiltinStyle(BuiltinStyleType.TITLE);

            Cell cell = workbook.getWorksheets().get(0).getCells().get("A1");
            cell.putValue("Aspose");
            cell.setStyle(style);

            workbook.getWorksheets().get(0).autoFitColumn(0);
            workbook.getWorksheets().get(0).autoFitRow(0);

            workbook.save(filePath + "UsingBuiltInStyles_Out.xlsx");
            workbook.save(filePath + "UsingBuiltInStyles_Out.ods");
        } catch (Exception e) {
            Log.e(TAG, "Using Built-in Styles", e);
        }
    }
}
