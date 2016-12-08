package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class RenderOnePDFPagePerExcelWorksheet {

    private static final String TAG = RenderOnePDFPagePerExcelWorksheet.class.getName();

    public void renderOnePDFPagePerExcelWorksheet() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Initialize a new Workbook
            //Open an Excel file
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");

            //Implement one page per worksheet option
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.setOnePagePerSheet(true);

            //Save the PDF file
            workbook.save(filePath + "RenderOnePDFPage_Out.pdf", pdfSaveOptions);

        } catch (Exception e) {
            Log.e(TAG, "Render One PDF Page Per Excel Worksheet", e);
        }
    }

}
