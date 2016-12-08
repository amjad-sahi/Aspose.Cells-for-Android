package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.Trendline;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class GetEquationTextOfChartTrendline {

    private static final String TAG = GetEquationTextOfChartTrendline.class.getName();

    public void getEquationTextOfChartTrendline() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object from source Excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access the first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the first chart inside the worksheet
            Chart chart = worksheet.getCharts().get(0);

            //Calculate the Chart first to get the Equation Text of Trendline
            chart.calculate();

            //Access the Trendline
            Trendline trendLine = chart.getNSeries().get(0).getTrendLines().get(0);

            //Read the Equation Text of Trendline
            Log.i(TAG, "Equation Text: " + trendLine.getDataLabels().getText());
        } catch (Exception e) {
            Log.e(TAG, "Get Equation Text of Chart Trendline", e);
        }
    }

}
