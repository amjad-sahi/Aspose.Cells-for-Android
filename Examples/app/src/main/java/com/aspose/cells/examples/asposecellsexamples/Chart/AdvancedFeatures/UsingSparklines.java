package com.aspose.cells.examples.asposecellsexamples.Chart.AdvancedFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellArea;
import com.aspose.cells.CellsColor;
import com.aspose.cells.Color;
import com.aspose.cells.Sparkline;
import com.aspose.cells.SparklineCollection;
import com.aspose.cells.SparklineGroup;
import com.aspose.cells.SparklineGroupCollection;
import com.aspose.cells.SparklineType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class UsingSparklines {

    private static final String TAG = UsingSparklines.class.getName();

    /**
     * Developers can create, delete or read sparklines (in the template file) using the API provided by Aspose.Cells.
     * By adding custom graphics for a given data range, developers have the freedom to add different types of tiny charts to selected cell areas.
     */
    public void usingSparklines() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiate a Workbook
            //Open a template file
            Workbook book = new Workbook(filePath + File.separator + "Book1.xlsx");
            //Get the first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Use the following lines if you need to read the Sparklines
            //Read the Sparklines from the template file (if it has)
            SparklineGroupCollection sgc = sheet.getSparklineGroupCollection();
            for(int i = 0; i < sgc.getCount(); i++ )
            {
                SparklineGroup g = sgc.get(i);
                //Display the Sparklines group information e.g type, number of sparklines items
                Log.v(TAG, "sparkline group: type:" + g.getType() + ", sparkline items count:" + g.getSparklineCollection().getCount());
                SparklineCollection sc = g.getSparklineCollection();
                for(int j = 0;j<sc.getCount(); j++)
                {
                    Sparkline s = sc.get(j);
                    //Display the individual Sparkines and the data ranges: Get where the sparkline is placed (Location range) and Data Range
                    Log.v(TAG, "sparkline: row:" + s.getRow() + ", col:" + s.getColumn() + ", dataRange:" + s.getDataRange());
                }
            }

            //Add new Sparklines
            //Define the CellArea D8:D10
            CellArea ca = CellArea.createCellArea(7, 3, 9, 3);

            //Add new Sparklines for a data range (A8:B10) to a cell area i.e. D8:D10
            int idx = sheet.getSparklineGroupCollection().add(SparklineType.COLUMN, "Sheet1!A8:B10", false, ca);
            SparklineGroup group = sheet.getSparklineGroupCollection().get(idx);
            //Create CellsColor
            CellsColor clr = book.createCellsColor();
            clr.setColor(Color.getOrange());
            group.setSeriesColor(clr);

            //Save the excel file
            book.save(filePath + File.separator + "UsingSparklines_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Using Sparklines", e);
        }
    }
}
