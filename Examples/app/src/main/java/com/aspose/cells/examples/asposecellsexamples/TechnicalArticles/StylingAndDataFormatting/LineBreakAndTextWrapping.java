package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class LineBreakAndTextWrapping {

    private static final String TAG = LineBreakAndTextWrapping.class.getName();

    public void wrappingText() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create Workbook Object
            Workbook wb = new Workbook();

            //Open first Worksheet in the workbook
            Worksheet ws = wb.getWorksheets().get(0);

            //Get Worksheet Cells Collection
            Cells cell = ws.getCells();

            //Increase the width of First Column Width
            cell.setColumnWidth(0, 35);

            //Increase the height of first row
            cell.setRowHeight(0, 65);

            //Add Text to the First Cell
            cell.get(0, 0).setValue("I am using the latest version of Aspose.Cells to test this functionality");

            //Get Cell's Style
            Style style = cell.get(0, 0).getStyle();

            //Set Text Wrap property to true
            style.setTextWrapped(true);

            //Set Cell's Style
            cell.get(0, 0).setStyle(style);

            //Save Excel File
            wb.save(filePath + "WrappingText_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Wrapping Text", e);
        }
    }

    public void explicitLineBreaks() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create Workbook Object
            Workbook wb = new Workbook();

            //Open first Worksheet in the workbook
            Worksheet ws = wb.getWorksheets().get(0);

            //Get Worksheet Cells Collection
            Cells cell = ws.getCells();

            //Increase the width of First Column Width
            cell.setColumnWidth(0, 35);

            //Increase the height of first row
            cell.setRowHeight(0, 65);

            // Add Text to the Firts Cell with Explicit Line Breaks
            cell.get(0, 0).setValue("I am using \nthe latest version of \nAspose.Cells \nto test this functionality");

            //Get Cell's Style
            Style style = cell.get(0, 0).getStyle();

            //Set Text Wrap property to true
            style.setTextWrapped(true);

            //Set Cell's Style
            cell.get(0, 0).setStyle(style);

            //Save Excel File
            wb.save(filePath + "ExplicitLineBreaks_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Explicit Line Breaks", e);
        }
    }
}
