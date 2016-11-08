package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.PaperSizeType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;

public class LoadWorkbookWithSpecifiedPrinterPaperSize {

    private static final String TAG = LoadWorkbookWithSpecifiedPrinterPaperSize.class.getName();

    public void loadWorkbookWithSpecifiedPrinterPaperSize() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a sample workbook and add some data inside the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.getWorksheets().get(0);
            worksheet.getCells().get("P30").putValue("This is sample data.");

            //Save the workbook in memory stream
            ByteArrayOutputStream baout = new ByteArrayOutputStream();
            workbook.save(baout, SaveFormat.XLSX);

            //Get bytes and create byte array input stream
            byte[] bts = baout.toByteArray();
            ByteArrayInputStream bain = new ByteArrayInputStream(bts);

            //Now load the workbook from memory stream with A5 paper size
            LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
            opts.setPaperSize(PaperSizeType.PAPER_A_5);
            workbook = new Workbook(bain, opts);

            //Save the workbook in pdf format
            workbook.save(filePath + "PrinterPaperSize_Out_a5.pdf");

            //Now load the workbook again from memory stream with A3 paper size
            opts = new LoadOptions(LoadFormat.XLSX);
            opts.setPaperSize(PaperSizeType.PAPER_A_3);
            workbook = new Workbook(bain, opts);

            //Save the workbook in pdf format
            workbook.save(filePath + "PrinterPaperSize_Out_a3.pdf");
        } catch (Exception e) {
            Log.e(TAG, "Load Workbook with specified Printer Paper Size", e);
        }
    }
}
