package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class AvoidExponentialNotationOfLargeNumbersWhileImportingFromHtml {

    private static final String TAG = AvoidExponentialNotationOfLargeNumbersWhileImportingFromHtml.class.getName();

    public void avoidExponentialNotationOfLargeNumbersWhileImportingFromHtml() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Sample Html containing large number with digits greater than 15
            String html = "<html>"
                    + "<body>"
                    + "<p>1234567890123456</p>"
                    + "</body>"
                    + "</html>";

            //Convert Html to byte array
            byte[] byteArray = html.getBytes();

            //Set Html load options and keep precision true
            HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
            loadOptions.setKeepPrecision(true);

            //Convert byte array into stream
            java.io.ByteArrayInputStream stream = new java.io.ByteArrayInputStream(byteArray);

            //Create workbook from stream with Html load options
            Workbook workbook = new Workbook(stream, loadOptions);

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Auto fit the sheet columns
            worksheet.autoFitColumns();

            //Save the workbook
            workbook.save(filePath + "AvoidExponentialNotation_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Avoid exponential notation of large numbers while importing from Html", e);
        }
    }
}
