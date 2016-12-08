package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class FitAllWorksheetColumnsOnSinglePDFPage {

    private static final String TAG = FitAllWorksheetColumnsOnSinglePDFPage.class.getName();

    public void fitWorksheetColumnsOnSinglePDFPage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create and initialize an instance of Workbook
            Workbook book = new Workbook(filePath + "QueryTable.xlsx");

            //Create and initialize an instance of PdfSaveOptions
            PdfSaveOptions saveOptions = new PdfSaveOptions(SaveFormat.PDF);

            //Set AllColumnsInOnePagePerSheet to true
            saveOptions.setAllColumnsInOnePagePerSheet(true);

            //Save Workbook to PDF fromart by passing the object of PdfSaveOptions
            book.save(filePath + "FitWorksheetColumns_Out.pdf", saveOptions);
        } catch (Exception e) {
            Log.e(TAG, "Fit Worksheet Columns on Single PDF Page", e);
        }
    }
}
