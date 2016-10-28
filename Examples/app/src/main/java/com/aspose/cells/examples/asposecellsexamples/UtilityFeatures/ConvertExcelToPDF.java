package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.DateTime;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.PdfCompliance;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class ConvertExcelToPDF {

    private static final String TAG = ConvertExcelToPDF.class.getName();

    public void convertExcelToPDF() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Save the document in PDF format
            workbook.save(filePath + File.separator + "ConvertExcelToPDF_Out.pdf", FileFormatType.PDF);
        } catch (Exception e) {
            Log.e(TAG, "Convert Excel to PDF", e);
        }
    }

    /**
     * Aspose.Cells for Java also provides support for PDF/A compliance.
     */
    public void pdfAConversion() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Define PdfSaveOptions
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            //Set the compliance type
            saveOptions.setCompliance(PdfCompliance.PDF_A_1_B);
            //Save the PDF file
            workbook.save(filePath + File.separator + "PDFAConversion_Out.pdf", saveOptions);
        } catch (Exception e) {
            Log.e(TAG, "PDF/A Conversion", e);
        }
    }

    public void setCreationTimeForOutputPDF() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Create an instance of PdfSaveOptions and pass SaveFormat to the constructor
            PdfSaveOptions options = new PdfSaveOptions(SaveFormat.PDF);

            //Set the CreatedTime for the PdfSaveOptions as per requirement
            options.setCreatedTime(DateTime.getNow());

            //Save the workbook to PDF format while passing the object of PdfSaveOptions
            workbook.save(filePath + File.separator + "CreationTimeForPDF_Out.pdf", SaveFormat.PDF);

        } catch (Exception e) {
            Log.e(TAG, "Set Creation Time For Output PDF", e);
        }
    }

}
