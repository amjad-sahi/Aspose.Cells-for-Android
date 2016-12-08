package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.util.Log;

import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AccessTextBoxByName {

    private static final String TAG = AccessTextBoxByName.class.getName();

    public void accessTextBoxByName() {
        try {
            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.getWorksheets().get(0);

            int idx = sheet.getTextBoxes().add(10, 10, 10, 10);

            //Create a texbox with some text and assign it some name
            TextBox tb1 = sheet.getTextBoxes().get(idx);
            tb1.setName("MyTextBox");
            tb1.setText("This is MyTextBox");

            //Access the same textbox via its name
            TextBox tb2 = sheet.getTextBoxes().get("MyTextBox");

            //Displaying the text of the textbox accessed by its name
            Log.i(TAG, tb2.getText());
        } catch (Exception e) {
            Log.e(TAG, "Get Validation Applied on a Cell", e);
        }
    }
}