package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class RemoveUnusedStylesInsideTheWorkbook {

    private static final String TAG = RemoveUnusedStylesInsideTheWorkbook.class.getName();

    public void removeUnusedStylesInsideTheWorkbook() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load template excel file containing unused styles
            Workbook workbook = new Workbook(filePath);

            //Remove all unused styles inside the template
            //this will also remove AsposeStyle which is an unused style inside the template
            workbook.removeUnusedStyles();

            //Save the file
            workbook.save(filePath + "RemoveUnusedStyle_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Remove unused styles inside the workbook", e);
        }
    }

}
