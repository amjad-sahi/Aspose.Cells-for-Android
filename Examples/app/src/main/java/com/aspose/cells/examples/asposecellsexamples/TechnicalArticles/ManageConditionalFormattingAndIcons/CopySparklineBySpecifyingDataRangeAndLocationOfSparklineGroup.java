package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageConditionalFormattingAndIcons;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.SparklineGroup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CopySparklineBySpecifyingDataRangeAndLocationOfSparklineGroup {

    private static final String TAG = CopySparklineBySpecifyingDataRangeAndLocationOfSparklineGroup.class.getName();

    public void copySparklineBySpecifyingDataRangeAndLocationOfSparklineGroup() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook from source Excel file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access the first sparkline group
            SparklineGroup group = worksheet.getSparklineGroupCollection().get(0);

            //Add Data Ranges and Locations inside this sparkline group
            group.getSparklineCollection().add("D5:O5", 4, 15);
            group.getSparklineCollection().add("D6:O6", 5, 15);
            group.getSparklineCollection().add("D7:O7", 6, 15);
            group.getSparklineCollection().add("D8:O8", 7, 15);

            //Save the workbook
            workbook.save(filePath + "CopySparkline_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Copy Sparkline by Specifying Data Range and Location of Sparkline Group", e);
        }
    }
}
