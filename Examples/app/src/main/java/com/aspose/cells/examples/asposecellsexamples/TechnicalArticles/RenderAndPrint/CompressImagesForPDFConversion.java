package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class CompressImagesForPDFConversion {

    private static final String TAG = CompressImagesForPDFConversion.class.getName();

    public void compressImagesForPDFConversion() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Initialize a new Workbook by loading an existing spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Create an instance of the PdfSaveOptions class
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Set Image's desired PPI & quality
            pdfSaveOptions.setImageResample(300, 70);

            //Save the spreadsheet in PDF format
            workbook.save(filePath + "CompressImagesForPDFConversion_Out.pdf", pdfSaveOptions);
        } catch (Exception e) {
            Log.e(TAG, "Compress Images for PDF Conversion", e);
        }
    }
}
