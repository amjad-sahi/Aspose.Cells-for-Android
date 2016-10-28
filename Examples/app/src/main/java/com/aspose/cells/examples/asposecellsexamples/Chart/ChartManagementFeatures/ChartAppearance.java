package com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Area;
import com.aspose.cells.Axis;
import com.aspose.cells.Cells;
import com.aspose.cells.CellsColor;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartArea;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartFrame;
import com.aspose.cells.ChartMarkerType;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.ChartPointCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.FillType;
import com.aspose.cells.Font;
import com.aspose.cells.Line;
import com.aspose.cells.LineType;
import com.aspose.cells.Series;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Title;
import com.aspose.cells.WeightType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class ChartAppearance {

    private static final String TAG = ChartAppearance.class.getName();

    public void setChartArea() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = worksheets.get(0);
            Cells cells = worksheet.getCells();
            //Adding a sample value to "A1" cell
            cells.get("A1").setValue(50);

            //Adding a sample value to "A2" cell
            cells.get("A2").setValue(100);

            //Adding a sample value to "A3" cell
            cells.get("A3").setValue(150);

            //Adding a sample value to "B1" cell
            cells.get("B1").setValue(60);

            //Adding a sample value to "B2" cell
            cells.get("B2").setValue(32);

            //Adding a sample value to "B3" cell
            cells.get("B3").setValue(50);

            //Adding a chart to the worksheet
            ChartCollection charts = worksheet.getCharts();

            //Accessing the instance of the newly added chart
            int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 15, 5);
            Chart chart = charts.get(chartIndex);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.add("A1:B3", true);

            //Setting the foreground color of the plot area
            ChartFrame plotArea = chart.getPlotArea();
            Area area = plotArea.getArea();
            area.setForegroundColor(Color.getBlue());

            //Setting the foreground color of the chart area
            ChartArea chartArea = chart.getChartArea();
            area = chartArea.getArea();
            area.setForegroundColor(Color.getYellow());

            //Setting the foreground color of the 1st NSeries area
            Series aSeries = nSeries.get(0);
            area = aSeries.getArea();
            area.setForegroundColor(Color.getRed());

            //Setting the foreground color of the area of the 1st NSeries point
            ChartPointCollection chartPoints = aSeries.getPoints();
            ChartPoint point = chartPoints.get(0);
            point.getArea().setForegroundColor(Color.getCyan());

            //Save the workbook
            workbook.save(filePath + File.separator + "ChartArea_Out.xls");

        } catch (Exception e) {
            Log.e(TAG, "Set Chart Area", e);
        }
    }

    public void setChartLines() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = worksheets.get(0);
            Cells cells = worksheet.getCells();
            //Adding a sample value to "A1" cell
            cells.get("A1").setValue(50);

            //Adding a sample value to "A2" cell
            cells.get("A2").setValue(100);

            //Adding a sample value to "A3" cell
            cells.get("A3").setValue(150);

            //Adding a sample value to "B1" cell
            cells.get("B1").setValue(60);

            //Adding a sample value to "B2" cell
            cells.get("B2").setValue(32);

            //Adding a sample value to "B3" cell
            cells.get("B3").setValue(50);

            //Adding a chart to the worksheet
            ChartCollection charts = worksheet.getCharts();

            //Accessing the instance of the newly added chart
            int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 15, 5);
            Chart chart = charts.get(chartIndex);

            //Adding NSeries (chart data source) to the chart ranging from "A1" cell
            SeriesCollection nSeries = chart.getNSeries();
            nSeries.add("A1:B3", true);

            //Applying a dotted line style on the lines of an NSeries
            Series aSeries = nSeries.get(0);
            Line line = aSeries.getLine();
            line.setStyle(LineType.DOT);

            //Applying a triangular marker style on the data markers of an NSeries
            aSeries.setMarkerStyle(ChartMarkerType.TRIANGLE);

            //Setting the weight of all lines in an NSeries to medium
            aSeries = nSeries.get(1);
            line = aSeries.getLine();
            line.setWeight(WeightType.MEDIUM_LINE);

        } catch (Exception e) {
            Log.e(TAG, "Set Chart Lines", e);
        }
    }

    public void applyMicrosoftExcel20072010ThemesToCharts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first chart in the sheet
            Chart chart = worksheet.getCharts().get(0);

            //Specify the FilFormat's type to Solid Fill of the first series
            chart.getNSeries().get(0).getArea().getFillFormat().setFillType(FillType.SOLID);

            //Get the CellsColor of SolidFill
            CellsColor cc = chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().getCellsColor();

            //Create a theme in Accent style
            cc.setThemeColor(new ThemeColor(ThemeColorType.ACCENT_6, 0.6));

            //Apply the them to the series
            chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().setCellsColor(cc);

            //Save the Excel file
            workbook.save(filePath + "ApplyMSExcelThemes_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Applying Microsoft Excel 2007/2010 Themes to Charts", e);
        }
    }

    public void setTitlesOfChartsOrAxes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first chart in the sheet
            Chart chart = worksheet.getCharts().get(0);

            //Setting the title of a chart
            Title title = chart.getTitle();
            title.setText("Title");

            //Setting the font color of the chart title to blue
            Font font = title.getTextFont();
            font.setColor(Color.getBlue());

            //Setting the title of category axis of the chart
            Axis categoryAxis = chart.getCategoryAxis();
            title = categoryAxis.getTitle();
            title.setText("Category");

            //Setting the title of value axis of the chart
            Axis valueAxis = chart.getValueAxis();
            title = valueAxis.getTitle();
            title.setText("Value");
        } catch (Exception e) {
            Log.e(TAG, "Setting the Titles of Charts or Axes", e);
        }
    }

    public void hidingMajorGridlines() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first chart in the sheet
            Chart chart = worksheet.getCharts().get(0);

            //Hiding the major gridlines of value axis
            Axis valueAxis = chart.getValueAxis();
            Line majorGridLines = valueAxis.getMajorGridLines();
            majorGridLines.setVisible(false);
        } catch (Exception e) {
            Log.e(TAG, "Hiding Major Gridlines", e);
        }
    }

    public void changingMajorGridlinesSettings() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first chart in the sheet
            Chart chart = worksheet.getCharts().get(0);

            //Setting the color of major gridlines of value axis to silver
            Axis categoryAxis = chart.getCategoryAxis();
            categoryAxis.getMajorGridLines().setColor(Color.getSilver());
        } catch (Exception e) {
            Log.e(TAG, "Changing Major Gridlines Settings", e);
        }
    }

    public void setBordersForBackAndSideWalls() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Get the first chart in the sheet
            Chart chart = worksheet.getCharts().get(0);

            //Get the side wall border line
            Line sideLine =  chart.getSideWall().getBorder();
            //Make it visible
            sideLine.setVisible(true);
            //Set the solid line
            sideLine.setStyle(LineType.SOLID);
            //Set the line width
            sideLine.setWeight(10);
            //Set the color
            sideLine.setColor(Color.getBlack());
        } catch (Exception e) {
            Log.e(TAG, "Set Borders for Back and Side Walls", e);
        }
    }
}
