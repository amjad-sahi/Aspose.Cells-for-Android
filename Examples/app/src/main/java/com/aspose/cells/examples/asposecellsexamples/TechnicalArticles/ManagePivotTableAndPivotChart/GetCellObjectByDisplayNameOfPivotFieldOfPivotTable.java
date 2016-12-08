package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManagePivotTableAndPivotChart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Color;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GetCellObjectByDisplayNameOfPivotFieldOfPivotTable {

    private static final String TAG = GetCellObjectByDisplayNameOfPivotFieldOfPivotTable.class.getName();

    public void getCellObjectByDisplayNameOfPivotFieldOfPivotTable() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first pivot table inside the worksheet
            PivotTable pivotTable = worksheet.getPivotTables().get(0);

            //Access cell by display name of 2nd data field of the pivot table
            String displayName = pivotTable.getDataFields().get(1).getDisplayName();
            Cell cell = pivotTable.getCellByDisplayName(displayName);

            //Access cell style and set its fill color and font color
            Style style = cell.getStyle();
            style.setForegroundColor(Color.getLightBlue());
            style.getFont().setColor(Color.getBlack());

            //Set the style of the cell
            pivotTable.format(cell.getRow(), cell.getColumn(), style);

            //Save workbook
            workbook.save(filePath + "CellObjectByDisplayName_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Get the Cell object by DisplayName of PivotField of PivotTable", e);
        }
    }
}
