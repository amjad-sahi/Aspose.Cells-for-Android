package com.aspose.cells.examples.asposecellsexamples.DrawingObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class Pictures {
    private static final String TAG = Pictures.class.getName();

    public void addPicture() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            WorksheetCollection worksheets = workbook.getWorksheets();

            //Obtaining the reference of first worksheet
            Worksheet sheet = worksheets.get(0);

            //Adding a picture at the location of a cell whose row and column indices
            //are 5 in the worksheet. It is "F6" cell
            int pictureIndex = sheet.getPictures().add(5, 5, filePath + File.separator + "school.jpg");
            Picture picture = sheet.getPictures().get(pictureIndex);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddPicture_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Pictures", e);
        }
    }

    /**
     * Position the picture with the Picture object's setLeftPositionInPixel and setTopPositionInPixel methods.
     */
    public void positionPicture() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the newly added worksheet.
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Adding a picture at the location of a cell whose row and column indices
            //are 5 in the worksheet. It is "F6" cell
            int pictureIndex = worksheet.getPictures().add(5, 5, filePath + File.separator + "school.jpg");
            Picture picture = worksheet.getPictures().get(pictureIndex);

            //Positioning the picture proportional to row height and colum width
            picture.setUpperDeltaX(200);
            picture.setUpperDeltaY(200);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "PositionPicture_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Position Picture", e);
        }
    }
}
