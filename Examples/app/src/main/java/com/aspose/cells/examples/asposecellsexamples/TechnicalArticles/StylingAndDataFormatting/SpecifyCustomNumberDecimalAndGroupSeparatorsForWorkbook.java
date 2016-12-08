package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SpecifyCustomNumberDecimalAndGroupSeparatorsForWorkbook {

    private static final String TAG = SpecifyCustomNumberDecimalAndGroupSeparatorsForWorkbook.class.getName();

    public void specifyCustomSeparators() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook workbook = new Workbook();

            //Specify custom separators
            workbook.getSettings().setNumberDecimalSeparator('.');
            workbook.getSettings().setNumberGroupSeparator(' ');

            //Test the custom separators
            //Get first worksheet from the collection
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get cell A1 and insert a new value
            Cell cell = worksheet.getCells().get("A1");
            cell.putValue(123456.789);

            //Get the style from cell A1 and set a new custom pattern
            Style style = cell.getStyle();
            style.setCustom("#,##0.000;[Red]#,##0.000");
            cell.setStyle(style);

            //AutoFit columns
            worksheet.autoFitColumns();

            //Save workbook in PDF format
            workbook.save(filePath + "SpecifyCustomSeparators_Out.pdf");
        } catch (Exception e) {
            Log.e(TAG, "Specifying Custom Separators", e);
        }
    }

}
