package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetDefaultFontWhileRenderingSpreadsheetToHTML {
    private static final String TAG = SetDefaultFontWhileRenderingSpreadsheetToHTML.class.getName();

    public void setDefaultFontWhileRenderingSpreadsheetToHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Access cell B4 and add some text to it
            Cell cell = sheet.getCells().get("B4");
            cell.putValue("This text has some unknown or invalid font which does not exist.");

            //Set the font of cell B4 which is unknown
            Style st = cell.getStyle();
            st.getFont().setName("UnknownNotExist");
            st.getFont().setSize(20);
            cell.setStyle(st);

            //Save the workbook in html format and set the default font to Courier New
            HtmlSaveOptions opts = new HtmlSaveOptions();
            opts.setDefaultFontName("Courier New");
            book.save(filePath + "Courier_New_Out.htm", opts);

            //Save the workbook in html format once again but set the default font to Arial
            opts.setDefaultFontName("Arial");
            book.save(filePath + "Arial_Out.htm", opts);

            //Save the workbook in html format once again but set the default font to Times New Roman
            opts.setDefaultFontName("Times New Roman");
            book.save(filePath + "Times_new_roman_Out.htm", opts);
        } catch (Exception e) {
            Log.e(TAG, "Set Default Font while Rendering Spreadsheet to HTML", e);
        }
    }
}
