package com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Series;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CustomChart {

    private static final String TAG = CustomChart.class.getName();

    /**
     * This example code below demonstrates how to create custom charts.
     * This example uses a column chart for the first data series and a line chart for the second series.
     * As a result, a column chart, combined with a line chart, is added to the worksheet.
     */
    public void createCustomChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the newly added worksheet
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
            Cells cells = worksheet.getCells();

            //Adding a sample value to "A1" cell
            cells.get("A1").setValue(50);

            //Adding a sample value to "A2" cell
            cells.get("A2").setValue(100);

            //Adding a sample value to "A3" cell
            cells.get("A3").setValue(150);

            //Adding a sample value to "A4" cell
            cells.get("A4").setValue(200);

            //Adding a sample value to "B1" cell
            cells.get("B1").setValue(60);

            //Adding a sample value to "B2" cell
            cells.get("B2").setValue(32);

            //Adding a sample value to "B3" cell
            cells.get("B3").setValue(50);

            //Adding a sample value to "B4" cell
            cells.get("B4").setValue(40);

            //Adding a chart to the worksheet
            ChartCollection charts = worksheet.getCharts();

            //Accessing the instance of the newly added chart
            int chartIndex = worksheet.getCharts().add(ChartType.COLUMN, 5, 0, 15, 5);
            Chart chart = worksheet.getCharts().get(chartIndex);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.add("A1:B4", true);

            //Setting the chart type of 2nd NSeries to display as line chart
            Series series = nSeries.get(1);
            series.setType(ChartType.LINE);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "CustomChart_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Create Custom Charts", e);
        }
    }
}
