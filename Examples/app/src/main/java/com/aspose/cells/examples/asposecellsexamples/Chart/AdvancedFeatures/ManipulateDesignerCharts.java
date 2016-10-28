package com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.CellsColor;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartArea;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.DataLabels;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.FillFormat;
import com.aspose.cells.FillType;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.LabelPositionType;
import com.aspose.cells.Legend;
import com.aspose.cells.LegendPositionType;
import com.aspose.cells.SheetType;
import com.aspose.cells.TextureType;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ManipulateDesignerCharts {
    private static final String TAG = ManipulateDesignerCharts.class.getName();

    /**
     * This example shows how to create a pie chart.
     */
    public void createAChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook();

            //Get the first worksheet.
            Worksheet sheet = workbook.getWorksheets().get(0);

            //Set the name of worksheet
            sheet.setName("Data");

            //Get the cells collection in the sheet.
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Put some values into a cells of the Data sheet.
            cells.get("A1").setValue("Region");
            cells.get("A2").setValue("France");
            cells.get("A3").setValue("Germany");
            cells.get("A4").setValue("England");
            cells.get("A5").setValue("Sweden");
            cells.get("A6").setValue("Italy");
            cells.get("A7").setValue("Spain");
            cells.get("A8").setValue("Portugal");
            cells.get("B1").setValue("Sale");
            cells.get("B2").setValue(70000);
            cells.get("B3").setValue(55000);
            cells.get("B4").setValue(30000);
            cells.get("B5").setValue(40000);
            cells.get("B6").setValue(35000);
            cells.get("B7").setValue(32000);
            cells.get("B8").setValue(10000);

            //Add a chart sheet.
            int sheetIndex = workbook.getWorksheets().add(SheetType.CHART);
            sheet = workbook.getWorksheets().get(sheetIndex);

            //Set the name of worksheet
            sheet.setName("Chart");

            //Create chart
            int chartIndex = sheet.getCharts().add(ChartType.PIE,1,1,25,10);
            Chart chart = sheet.getCharts().get(chartIndex);

            //Set some properties of chart plot area.
            //to set the fill color and make the border invisible.
            chart.getPlotArea().getArea().setForegroundColor(Color.getCyan());
            chart.getPlotArea().getArea().getFillFormat().setTwoColorGradient(Color.getYellow(), Color.getWhite(), GradientStyleType.VERTICAL, 2);
            chart.getPlotArea().getBorder().setVisible(false);

            //Set properties of chart title
            chart.getTitle().setText("Sales By Region");
            chart.getTitle().getTextFont().setColor(Color.getBlue());
            chart.getTitle().getTextFont().setBold(true);
            chart.getTitle().getTextFont().setSize(12);

            //Set properties of nseries
            chart.getNSeries().add("Data!B2:B8", true);
            chart.getNSeries().setCategoryData("Data!A2:A8");

            //Set the DataLabels in the chart
            DataLabels datalabels = null;
            int i = 0;
            for (i = 0; i < chart.getNSeries().getCount(); i++) {
                datalabels = chart.getNSeries().get(i).getDataLabels();
                datalabels.setPosition(LabelPositionType.OUTSIDE_END);
                datalabels.setShowCategoryName(true);
                datalabels.setShowValue(true);
                datalabels.setShowPercentage(false);
                datalabels.setShowLegendKey(false);
            }

            //Set the ChartArea.
            ChartArea chartarea = chart.getChartArea();
            chartarea.getArea().getFillFormat().setTexture(TextureType.BLUE_TISSUE_PAPER);

            //Set the Legend.
            Legend legend = chart.getLegend();
            legend.setPosition(LegendPositionType.LEFT);
            legend.setHeight(600);
            legend.setWidth(700);
            legend.getTextFont().setBold(true);
            legend.getBorder().setColor(Color.getBlue());
            //Set FillFormat.
            FillFormat fillformat = legend.getArea().getFillFormat();
            fillformat.setTexture(TextureType.BOUQUET);

            //Save the Excel file
            workbook.save(filePath + File.separator + "CreateAChart_Out.xls", FileFormatType.EXCEL_97_TO_2003);
        } catch(Exception e) {
            Log.e(TAG, "Creating a Chart", e);
        }
    }

    /**
     * The following example shows how to manipulate an existing chart by changing the Label "England, 30000" to "United Kingdom, 30K".
     */
    public void manipulateTheChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook(filePath + File.separator + "PieChart.xls");

            //Get the designer chart in the second sheet.
            Worksheet sheet = workbook.getWorksheets().get(1);
            Chart chart = sheet.getCharts().get(0);

            //Get the data labels in the data series of the third data point.
            DataLabels datalabels = chart.getNSeries().get(0).getPoints().get(2).getDataLabels();

            //Change the text of the label.
            datalabels.setText("United Kingdom, 30K ");

            //Save the excel file.
            workbook.save(filePath + File.separator + "ManipulateTheChart_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Manipulating the Chart", e);
        }
    }

    /**
     * This example manipulates a line chart by adding two data series and changing the line colors.
     */
    public void manipulateALineChartInTheDesignerTemplate() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a new Workbook.
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Get the designer chart in the first worksheet.
            Chart chart = workbook.getWorksheets().get(0).getCharts().get(0);

            //Add the third data series to it.
            chart.getNSeries().add("{60, 80, 10}", true);

            //Add another data series (fourth) to it.
            chart.getNSeries().add("{0.3, 0.7, 1.2}", true);

            //Plot the fourth data series on the second axis.
            chart.getNSeries().get(3).setPlotOnSecondAxis(true);

            //Change the line color of the second data series.
            chart.getNSeries().get(1).getLine().setColor(Color.getGreen());

            //Change the line color of the third data series.
            chart.getNSeries().get(2).getLine().setColor(Color.getRed());

            //Make the second value axis visible.
            chart.getSecondValueAxis().setVisible(true);

            //Save the excel file.
            workbook.save(filePath + File.separator + "ManipulateLinechart_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Manipulating a Line Chart in the Designer Template", e);
        }
    }

    /**
     * Apply different Microsoft Excel themes and colors to the SeriesCollection collection or other chart objects.
     */
    public void applyMicrosoftExcel20072010ThemesToCharts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate the workbook to open the file that contains a chart
            Workbook workbook = new Workbook(filePath + File.separator + "source.xlsx");

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
            workbook.save(filePath + "MSExcel20072010Themes_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Applying Microsoft Excel 2007/2010 Themes to Charts", e);
        }
    }
}
