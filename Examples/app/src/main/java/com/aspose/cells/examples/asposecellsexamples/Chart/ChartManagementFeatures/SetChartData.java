package com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class SetChartData {

    private static final String TAG = SetChartData.class.getName();

    /** Chart data is the data used as a data source for charts.
     * Add a range of cells that contain the chart data by calling the SeriesCollection object's Add method.
     */
    public void chartData() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Chart chart = worksheet.getCharts().get(0);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.add("A1:B4", true);

        } catch(Exception e) {
            Log.e(TAG, "Chart Data", e);
        }
    }

    /** Category data is used to label chart data and
     * can be added to the SeriesCollection collection by using its setCategoryData method.
     */
    public void categoryData() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");
            Worksheet worksheet = workbook.getWorksheets().get(0);
            Chart chart = worksheet.getCharts().get(0);

            //Setting the data source for the category data of NSeries
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.setCategoryData("C1:C4");

        } catch(Exception e) {
            Log.e(TAG, "Category Data", e);
        }
    }

    /**
     * This example demonstrate the use of chart and category data.
     * Executing the example code adds a column chart to the worksheet.
     */
    public void setChartAndCategoryData() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Obtaining the reference of the first worksheet
            Worksheet worksheet= worksheets.get(0);
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

            //Adding a sample value to "C1" cell as category data
            cells.get("C1").setValue("Q1");

            //Adding a sample value to "C2" cell as category data
            cells.get("C2").setValue("Q2");

            //Adding a sample value to "C3" cell as category data
            cells.get("C3").setValue("Y1");

            //Adding a sample value to "C4" cell as category data
            cells.get("C4").setValue("Y2");

            //Adding a chart to the worksheet
            ChartCollection charts = worksheet.getCharts();

            //Accessing the instance of the newly added chart
            int chartIndex = charts.add(ChartType.COLUMN,5,0,15,5);
            Chart chart = charts.get(chartIndex);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.add("A1:B4",true);

            //Setting the data source for the category data of NSeries
            nSeries.setCategoryData("C1:C4");

            workbook.save(filePath + File.separator + "ChartAndCategoryData_Out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Chart and Category Data", e);
        }
    }
}
