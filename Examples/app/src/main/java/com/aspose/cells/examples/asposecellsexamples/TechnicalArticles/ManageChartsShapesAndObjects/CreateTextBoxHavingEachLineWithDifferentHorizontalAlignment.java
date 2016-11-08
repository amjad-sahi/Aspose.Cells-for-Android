package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.TextParagraph;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class CreateTextBoxHavingEachLineWithDifferentHorizontalAlignment {

    private static final String TAG = CreateTextBoxHavingEachLineWithDifferentHorizontalAlignment.class.getName();

    public void createTextBoxHavingEachLineWithDifferentHorizontalAlignment() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Add text box to the sheet
            sheet.getShapes().addShape(MsoDrawingType.TEXT_BOX, 2, 0, 2, 0, 80, 400);

            //Access first shape which is a text box and set its text
            Shape shape = sheet.getShapes().get(0);
            shape.setText("Sign up for your free phone number.\nCall and text online for free.\nCall your friends and family.");

            //Access the first paragraph and set its horizontal alignment to left
            TextParagraph para = shape.getTextBody().getTextParagraphs().get(0);
            para.setAlignmentType(TextAlignmentType.LEFT);

            //Access the second paragraph and set its horizontal alignment to center
            para = shape.getTextBody().getTextParagraphs().get(1);
            para.setAlignmentType(TextAlignmentType.CENTER);

            //Access the third paragraph and set its horizontal alignment to right
            para = shape.getTextBody().getTextParagraphs().get(2);
            para.setAlignmentType(TextAlignmentType.RIGHT);

            //Save the result in XLSX format
            book.save(filePath + "CreateTextBox_Out.xlsx", SaveFormat.XLSX);

        } catch (Exception e) {
            Log.e(TAG, "Create TextBox Having Each Line with Different Horizontal Alignment", e);
        }
    }
}
