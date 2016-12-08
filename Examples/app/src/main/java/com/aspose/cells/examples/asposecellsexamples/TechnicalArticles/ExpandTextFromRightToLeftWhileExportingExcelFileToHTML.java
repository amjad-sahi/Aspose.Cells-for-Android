package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class ExpandTextFromRightToLeftWhileExportingExcelFileToHTML {

    private static final String TAG = ExpandTextFromRightToLeftWhileExportingExcelFileToHTML.class.getName();

    public void expandTextFromRightToLeftWhileExportingExcelFileToHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load source excel file inside the workbook object
            Workbook wb = new Workbook(filePath + "ExpandText.xlsx");

            //Save workbook in HTML format
            wb.save(filePath + "ExpandTextFromRightToLeft_Out_" + CellsHelper.getVersion() + ".html", SaveFormat.HTML);

        } catch (Exception e) {
            Log.e(TAG, "Expanding text from right to left while exporting Excel file to HTML", e);
        }
    }
}