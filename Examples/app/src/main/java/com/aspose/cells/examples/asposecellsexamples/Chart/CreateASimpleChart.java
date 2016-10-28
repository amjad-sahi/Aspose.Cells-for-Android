package com.aspose.cells.examples.asposecellsexamples.Chart;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class CreateASimpleChart {

    private static final String TAG = CreateASimpleChart.class.getName();

    /**
     * This example adds a pyramid chart to a spreadsheet.
     */
    public void createAChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet = worksheets.get(0);

            //Adding some sample value to cells
            Cells cells = sheet.getCells();
            Cell cell = cells.get("A1");
            cell.setValue(50);
            cell = cells.get("A2");
            cell.setValue(100);
            cell = cells.get("A3");
            cell.setValue(150);
            cell = cells.get("B1");
            cell.setValue(4);
            cell = cells.get("B2");
            cell.setValue(20);
            cell = cells.get("B3");
            cell.setValue(50);

            ChartCollection charts = sheet.getCharts();

            //Adding a chart to the worksheet
            int chartIndex = charts.add(ChartType.PYRAMID, 5, 0, 15, 5);
            Chart chart = charts.get(chartIndex);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B3"
            SeriesCollection series = chart.getNSeries();
            series.add("A1:B3", true);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "CreateAChart_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Creating a Simple Chart", e);
        }
    }
}
