package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DeleteRedundantSpacesAfterLineBreakWhileImportingHtml {

    private static final String TAG = DeleteRedundantSpacesAfterLineBreakWhileImportingHtml.class.getName();

    public void deleteRedundantSpacesAfterLineBreakWhileImportingHtml() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Sample Html containing redundant spaces after <br> tag
            String html = "<html>"
                    + "<body>"
                    + "<table>"
                    + "<tr>"
                    + "<td>"
                    + "<br>    This is sample data"
                    + "<br>    This is sample data"
                    + "<br>    This is sample data"
                    + "</td>"
                    + "</tr>"
                    + "</table>"
                    + "</body>"
                    + "</html>";

            //Convert Html to byte array
            byte[] byteArray = html.getBytes();

            //Set Html load options and keep precision true
            HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
            loadOptions.setDeleteRedundantSpaces(true);

            //Convert byte array into stream
            java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

            //Create workbook from stream with Html load options
            Workbook workbook = new Workbook(stream, loadOptions);

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto fit the sheet columns
            worksheet.autoFitColumns();

            //Save the workbook
            workbook.save(filePath + "output-" + loadOptions.getDeleteRedundantSpaces() + ".xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Delete redundant spaces after line break while importing Html", e);
        }
    }
}