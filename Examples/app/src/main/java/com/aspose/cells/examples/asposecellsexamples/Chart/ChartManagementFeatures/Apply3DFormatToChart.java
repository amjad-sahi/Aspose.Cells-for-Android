package com.aspose.cells.examples.asposecellsexamples.Chart.ChartManagementFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Bevel;
import com.aspose.cells.BevelPresetType;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.Format3D;
import com.aspose.cells.LightRigType;
import com.aspose.cells.PresetMaterialType;
import com.aspose.cells.Series;
import com.aspose.cells.ShapePropertyCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class Apply3DFormatToChart {

    private static final String TAG = Apply3DFormatToChart.class.getName();

    /**
     * This example demonstrate how to create a chart and apply Microsoft Excel 2007 3D formatting.
     * After executing the example code, a column chart (with 3D effects) will be added to the worksheet.
     */
    public void set3DFormatToChart() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a new Workbook
            Workbook book = new Workbook();
            //Add a Data Worksheet
            Worksheet dataSheet = book.getWorksheets().add("DataSheet");
            //Add Chart Worksheet
            Worksheet sheet = book.getWorksheets().add("MyChart");
            //Put some values into the cells in the data worksheet
            dataSheet.getCells().get("B1").putValue(1);
            dataSheet.getCells().get("B2").putValue(2);
            dataSheet.getCells().get("B3").putValue(3);
            dataSheet.getCells().get("A1").putValue("A");
            dataSheet.getCells().get("A2").putValue("B");
            dataSheet.getCells().get("A3").putValue("C");

            //Define the Chart Collection
            ChartCollection charts = sheet.getCharts();
            //Add a Column chart to the Chart Worksheet
            int chartSheetIdx = charts.add(ChartType.COLUMN, 5, 0, 25, 15);

            //Get the newly added Chart
            Chart chart = book.getWorksheets().get(2).getCharts().get(0);

            //Set the background/foreground color for PlotArea/ChartArea
            chart.getPlotArea().getArea().setBackgroundColor(Color.getWhite());
            chart.getChartArea().getArea().setBackgroundColor(Color.getWhite());
            chart.getPlotArea().getArea().setForegroundColor(Color.getWhite());
            chart.getChartArea().getArea().setForegroundColor(Color.getWhite());

            //Hide the Legend
            chart.setShowLegend(false);

            //Add Data Series for the Chart
            chart.getNSeries().add("DataSheet!B1:B3", true);
            //Specify the Category Data
            chart.getNSeries().setCategoryData("DataSheet!A1:A3");

            //Get the Data Series
            Series ser = chart.getNSeries().get(0);

            //Apply the 3D formatting
            ShapePropertyCollection spPr = ser.getShapeProperties();
            Format3D fmt3d = spPr.getFormat3D();
            //Specify Bevel with its height/width
            Bevel bevel = fmt3d.getTopBevel();
            bevel.setType(BevelPresetType.CIRCLE);
            bevel.setHeight(5);
            bevel.setWidth(9);
            //Specify Surface material type
            fmt3d.setSurfaceMaterialType(PresetMaterialType.WARM_MATTE);
            //Specify surface lighting type
            fmt3d.setSurfaceLightingType(LightRigType.THREE_POINT);
            //Specify lighting angle
            fmt3d.setLightingAngle(20);

            //Specify Series background/foreground and line color
            ser.getArea().setBackgroundColor(Color.getMaroon());
            ser.getArea().setForegroundColor(Color.getMaroon());
            ser.getBorder().setColor(Color.getMaroon());

            //Save the Excel file
            book.save(filePath + File.separator + "3DFormatToChart_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Set 3D Format to Chart", e);
        }
    }
}
