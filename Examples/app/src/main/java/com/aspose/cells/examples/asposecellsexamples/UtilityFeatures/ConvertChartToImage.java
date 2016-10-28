package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SeriesCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;
import java.io.FileOutputStream;

public class ConvertChartToImage {

    private static final String TAG = ConvertChartToImage.class.getName();

    public void convertChartToImage() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet =  worksheets.get(0);

            //Adding some sample value to cells
            Cells cells = sheet.getCells();
            Cell cell = cells.get("A1");
            cell.setValue(50);
            cell = cells.get("A2");
            cell. setValue (100);
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

            //Get the Chart image
            ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
            imgOpts.setImageFormat(ImageFormat.getEmf());

            //Save the chart image.
            chart.toImage(new FileOutputStream(filePath + File.separator + "ConvertChartToImage_Out.emf"), imgOpts);

        } catch (Exception e) {
            Log.e(TAG, "Convert Chart to Image", e);
        }
    }
}
