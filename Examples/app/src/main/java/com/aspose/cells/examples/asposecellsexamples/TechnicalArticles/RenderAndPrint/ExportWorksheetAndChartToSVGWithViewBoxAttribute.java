package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Chart;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExportWorksheetAndChartToSVGWithViewBoxAttribute {

    private static final String TAG = ExportWorksheetAndChartToSVGWithViewBoxAttribute.class.getName();

    public void exportWorksheetAndChartToSVGWithViewBoxAttribute() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of ImageOrPrintOptions class
            //Set SaveFormat to SVG & SVGFitToViewPort to true
            ImageOrPrintOptions opts = new ImageOrPrintOptions();
            opts.setSaveFormat(SaveFormat.SVG);
            opts.setSVGFitToViewPort(true);

            //Create an instance of Workbook and load an existing spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xlsx");
            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);
            //Access first chart from the worksheet
            Chart chart = worksheet.getCharts().get(0);
            //Save the chart to SVG format
            chart.toImage(filePath + "Chart_Out.svg", opts);

            //Create an instance of SheetRender while passing instances of Worksheet and ImageOrPrintOptions
            SheetRender render = new SheetRender(worksheet, opts);
            //Save the worksheet to SVG format
            render.toImage(0, filePath + "Sheet_Out.svg");
        } catch (Exception e) {
            Log.e(TAG, "Export Worksheet and Chart to SVG with ViewBox Attribute", e);
        }
    }

}
