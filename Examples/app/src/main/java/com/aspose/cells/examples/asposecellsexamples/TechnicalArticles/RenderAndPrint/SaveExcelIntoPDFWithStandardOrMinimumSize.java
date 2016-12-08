package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PdfOptimizationType;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class SaveExcelIntoPDFWithStandardOrMinimumSize {

    private static final String TAG = SaveExcelIntoPDFWithStandardOrMinimumSize.class.getName();

    public void saveExcelIntoPDFWithStandardOrMinimumSize() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Load excel file into workbook object
            Workbook workbook = new Workbook(filePath + "sample.xlsx");

            //Save into Pdf with Minimum size
            PdfSaveOptions opts = new PdfSaveOptions();
            opts.setOptimizationType(PdfOptimizationType.MINIMUM_SIZE);

            workbook.save(filePath + "SaveExcelIntoPDF_Out.pdf", opts);
        } catch (Exception e) {
            Log.e(TAG, "Save Excel into PDF with Standard or Minimum Size", e);
        }
    }
}
