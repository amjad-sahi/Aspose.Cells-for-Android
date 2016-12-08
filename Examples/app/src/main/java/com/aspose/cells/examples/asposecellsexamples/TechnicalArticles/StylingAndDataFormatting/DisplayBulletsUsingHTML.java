package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class DisplayBulletsUsingHTML {

    private static final String TAG = DisplayBulletsUsingHTML.class.getName();

    public void displayBulletsUsingHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create a workbook object
            Workbook workbook = new Workbook();

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access cell A1
            Cell cell = worksheet.getCells().get("A1");

            //Set the HTML string
            cell.setHtmlString("<font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'>Text 1 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 2 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 3 </font><font style='font-family:Wingdings;font-size:8.0pt;color:#009DD9;mso-font-charset:2;'>l</font><font style='font-family:Arial;font-size:10pt;color:#666666;vertical-align:top;text-align:left;'> Text 4 </font>");

            //Save the spreadsheet
            workbook.save(filePath + "DisplayBulletsUsingHTML_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Display Bullets using HTML", e);
        }
    }

}
