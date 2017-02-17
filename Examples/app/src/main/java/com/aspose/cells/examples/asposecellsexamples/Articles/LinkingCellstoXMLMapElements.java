package com.aspose.cells.examples.asposecellsexamples.Articles;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.*;
import com.aspose.cells.examples.asposecellsexamples.MainActivity;

import java.io.File;
import java.io.InputStream;

public class LinkingCellstoXMLMapElements {

    private static final String TAG = "Aspose.Cells.Examples";

    /**
     * Run Code: Linking Cells to XML Map Elements
     */
    public static void Run(Context context)
    {
        Log.w(TAG, "Running LinkingCellstoXMLMapElements");

        try
        {
            String SD_PATH = Environment.getExternalStorageDirectory().toString() + "/Aspose/";

            //Read the sample workbook from assest
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open("sampleLinkingCellstoXMLMapElements.xlsx");

            //Load a sample spreadsheet
            Workbook book = new Workbook(in);

            //Access the XML Map from the spreadsheet
            XmlMap map = book.getWorksheets().getXmlMaps().get(0);

            //Access first worksheet from the collection
            Worksheet sheet = book.getWorksheets().get(0);

            //Map FIELD1 and FIELD2 to cell A1 and B2
            sheet.getCells().linkToXmlMap(map.getName(), 0, 0, "/root/row/FIELD1");
            sheet.getCells().linkToXmlMap(map.getName(), 1, 1, "/root/row/FIELD2");

            //Map FIELD4 and FIELD5 to cell C3 and D4
            sheet.getCells().linkToXmlMap(map.getName(), 2, 2, "/root/row/FIELD4");
            sheet.getCells().linkToXmlMap(map.getName(), 3, 3, "/root/row/FIELD5");

            //Map FIELD7 and FIELD8 to cell E5 and F6
            sheet.getCells().linkToXmlMap(map.getName(), 4, 4, "/root/row/FIELD7");
            sheet.getCells().linkToXmlMap(map.getName(), 5, 5, "/root/row/FIELD8");

            //Save the result in XLSX format
            book.save(SD_PATH + "outputLinkingCellstoXMLMapElements.xlsx");

            Log.w(TAG, "outputLinkingCellstoXMLMapElements.xlsx created successfully");
        }
        catch (Exception ex)
        {
            Log.e(TAG, "Some exception occurred in LinkingCellstoXMLMapElements");
            Log.e(TAG, "Exception: " + ex.getMessage());
            Log.e(TAG, "StackTrace: " + Log.getStackTraceString(ex));
        }

    }
}
