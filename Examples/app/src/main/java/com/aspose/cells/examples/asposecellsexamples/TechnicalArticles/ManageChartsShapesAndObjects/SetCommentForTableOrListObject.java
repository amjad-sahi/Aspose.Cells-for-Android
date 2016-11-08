package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.ListObject;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetCommentForTableOrListObject {
    private static final String TAG = SetCommentForTableOrListObject.class.getName();

    /**
     * You can set the comments for an Excel Table or List Object inside the worksheet using the ListObject.Comment property.
     * The comment will be visible inside the xl/tables/tableName.xml file.
     */
    public void setCommentForTableOrListObject() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Open the template file
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access first list object or table
            ListObject lstObj = worksheet.getListObjects().get(0);

            //Set the comment of the list object
            lstObj.setComment("This is Aspose.Cells comment.");

            //Save the workbook in XLSX format
            workbook.save(filePath + "SetCommentForTableOrListObject_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Set the Comment for Table or List Object inside the Worksheet", e);
        }
    }
}
