package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class LimitTheNumberOfPagesGeneratedExcelToPDFConversion {

    private static final String TAG = LimitTheNumberOfPagesGeneratedExcelToPDFConversion.class.getName();

    public void limitTheNumberOfPagesGenerated() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Open an Excel file
            Workbook wb = new Workbook(filePath + "Book1.xlsx");
            //Instantiate the PdfSaveOption
            PdfSaveOptions options = new PdfSaveOptions();

            //Print only Page 3 and Page 4 in the output PDF
            //Starting page index (0-based index)
            options.setPageIndex(2);
            //Number of pages to be printed
            options.setPageCount(2);

            //Save the PDF file
            wb.save(filePath + "RenderARangeOfPages_Out.pdf", options);
        } catch (Exception e) {
            Log.e(TAG, "Limit the Number of Pages Generated - Excel to PDF Conversion", e);
        }
    }
}
