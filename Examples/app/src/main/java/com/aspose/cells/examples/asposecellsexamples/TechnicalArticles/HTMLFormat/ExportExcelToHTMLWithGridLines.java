package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExportExcelToHTMLWithGridLines {

    private static final String TAG = ExportExcelToHTMLWithGridLines.class.getName();

    /**
     * Export Excel to HTML with GridLines
     */
    public void exportExcelToHTMLWithGridLines() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Fill worksheet with some integer values
            for (int r = 0; r < 10; r++ ) {
                for(int c = 0; c < 10; c++) {
                    sheet.getCells().get(r, c).putValue(r*1);
                }
            }

            //Save result in HTML format and export Grid Lines
            HtmlSaveOptions opts = new HtmlSaveOptions();
            opts.setExportGridLines(true);
            book.save(filePath + "ExportExcelToHTMLWithGridLines_Out.html", opts);
        } catch (Exception e) {
            Log.e(TAG, "Export Excel to HTML with GridLines", e);
        }
    }
}
