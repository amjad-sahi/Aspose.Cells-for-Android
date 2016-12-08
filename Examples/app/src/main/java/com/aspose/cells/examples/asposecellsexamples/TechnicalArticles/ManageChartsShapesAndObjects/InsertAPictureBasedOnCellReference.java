package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cells;
import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;

import java.io.File;
import java.io.FileInputStream;

public class InsertAPictureBasedOnCellReference {

    private static final String TAG = InsertAPictureBasedOnCellReference.class.getName();

    public void insertAPictureBasedOnCellReference() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new Workbook
            Workbook workbook = new Workbook();

            //Get the first worksheet's cells collection
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Add string values to the cells
            cells.get("A1").putValue("A1");
            cells.get("C10").putValue("C10");

            //Add a picture to the D1 cell
            FileInputStream stream = new FileInputStream(filePath + "footer.jpg");
            Picture pic = (Picture)workbook.getWorksheets().get(0).getShapes().addPicture(0, 3, stream, 10, 10);

            //Set the size of the picture.
            pic.setHeightCM(4.48);
            pic.setWidthCM(5.28);

            //Specify the formula that refers to the source range of cells
            pic.setFormula("A1:C10");

            //Update the shapes selected value in the worksheet
            workbook.getWorksheets().get(0).getShapes().updateSelectedValue();

            //Save the Excel file.
            workbook.save(filePath + "PictureBasedOnCellReference_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Insert a Picture based on Cell Reference", e);
        }
    }

}
