package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ConvertChartToImageInSVGFormat {

    private static final String TAG = ConvertChartToImageInSVGFormat.class.getName();

    public void convertChartToImageInSVGFormat() {
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

            //Save the chart into image in SVG format
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.setSaveFormat(SaveFormat.SVG);
            chart.toImage(filePath + "ChartImage_Out.svg", options);
        } catch (Exception e) {
            Log.e(TAG, "Convert Chart to Image in SVG Format", e);
        }
    }
}
