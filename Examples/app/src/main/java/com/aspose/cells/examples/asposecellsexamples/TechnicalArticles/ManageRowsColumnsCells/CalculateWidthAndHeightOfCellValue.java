package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageRowsColumnsCells;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CalculateWidthAndHeightOfCellValue {

    private static final String TAG = CalculateWidthAndHeightOfCellValue.class.getName();

    public void calculateWidthAndHeightOfCellValueInUnitOfPixels() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook();

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access cell B2 and add some value inside it
            Cell cell = worksheet.getCells().get("B2");
            cell.putValue("Welcome to Aspose!");

            //Enlarge its font to size 16
            Style style = cell.getStyle();
            style.getFont().setSize(16);
            cell.setStyle(style);

            //Calculate the width and height of the cell value in unit of pixels
            int widthOfValue = cell.getWidthOfValue();
            int heightOfValue = cell.getHeightOfValue();

            //Print both values
            Log.i(TAG, "Width of Cell Value: " + widthOfValue);
            Log.i(TAG, "Height of Cell Value: " + heightOfValue);

            //Set the row height and column width to adjust/fit the cell value inside cell
            worksheet.getCells().setColumnWidthPixel(1, widthOfValue);
            worksheet.getCells().setRowHeightPixel(1, heightOfValue);

            //Save the output excel file
            workbook.save(filePath + "CalculateWidthAndHeight_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Calculate the Width and Height of the Cell Value in Unit of Pixels", e);
        }
    }
}
