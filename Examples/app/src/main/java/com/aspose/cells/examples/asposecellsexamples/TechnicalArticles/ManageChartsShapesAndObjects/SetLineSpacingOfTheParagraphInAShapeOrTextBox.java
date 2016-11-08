package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageChartsShapesAndObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.LineSpaceSizeType;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.TextParagraph;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class SetLineSpacingOfTheParagraphInAShapeOrTextBox {

    private static final String TAG = SetLineSpacingOfTheParagraphInAShapeOrTextBox.class.getName();

    public void setLineSpacingOfTheParagraphInAShapeOrTextBox() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create an instance of Workbook
            Workbook book = new Workbook();

            //Access first worksheet
            Worksheet sheet = book.getWorksheets().get(0);

            //Add text box to the sheet
            sheet.getShapes().addShape(MsoDrawingType.TEXT_BOX, 2, 0, 2, 0, 100, 200);

            //Access first shape which is a text box and set its text
            Shape shape = sheet.getShapes().get(0);
            shape.setText("Sign up for your free phone number.\nCall and text online for free.");

            //Access the first paragraph
            TextParagraph para = shape.getTextBody().getTextParagraphs().get(1);

            //Set the line space
            para.setLineSpaceSizeType(LineSpaceSizeType.POINTS);
            para.setLineSpace(20);

            //Set the space after
            para.setSpaceAfterSizeType(LineSpaceSizeType.POINTS);
            para.setSpaceAfter(10);

            //Set the space before
            para.setSpaceBeforeSizeType(LineSpaceSizeType.POINTS);
            para.setSpaceBefore(10);

            //Save the workbook in xlsx format
            book.save(filePath + "SetLineSpacing_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Set Line Spacing of the Paragraph in a Shape or TextBox", e);
        }
    }
}
