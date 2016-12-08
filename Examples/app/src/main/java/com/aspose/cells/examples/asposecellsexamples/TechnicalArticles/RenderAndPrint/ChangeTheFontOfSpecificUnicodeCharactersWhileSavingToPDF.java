package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ChangeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF {

    private static final String TAG = ChangeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF.class.getName();

    public void changeTheFontOfSpecificUnicodeCharactersWhileSavingToPDF() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook();

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access cells
            Cell cell1 = worksheet.getCells().get("A1");
            Cell cell2 = worksheet.getCells().get("B1");

            //Set the styles of both cells to Times New Roman
            Style style = cell1.getStyle();
            style.getFont().setName("Times New Roman");
            cell1.setStyle(style);
            cell2.setStyle(style);

            //Put the values inside the cell
            cell1.putValue("Hello without Non-Breaking Hyphen");
            cell2.putValue("Hello" + (char) (8209) + " with Non-Breaking Hyphen");

            //Autofit the columns
            worksheet.autoFitColumns();

            //Save PDF without setting PdfSaveOptions.IsFontSubstitutionCharGranularity
            workbook.save(filePath + "ChangeTheFontOfSpecificUnicodeCharacters_out1.pdf");

            //Save PDF after setting PdfSaveOptions.IsFontSubstitutionCharGranularity to true
            PdfSaveOptions opts = new PdfSaveOptions();
            opts.setFontSubstitutionCharGranularity(true);
            workbook.save(filePath + "ChangeTheFontOfSpecificUnicodeCharacters_out2.pdf", opts);

        } catch (Exception e) {
            Log.e(TAG, "Change the Font of Specific Unicode Characters while Saving to PDF", e);
        }
    }
}
