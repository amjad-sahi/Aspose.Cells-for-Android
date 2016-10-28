package com.aspose.cells.examples.asposecellsexamples.DrawingObjects;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Comment;
import com.aspose.cells.Font;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class Comments {
    private static final String TAG = Comments.class.getName();

    /**
     * Add a comment to a cell by calling the Shapes collection's addComments method (encapsulated in the Worksheet object).
     * The new Comment object can be accessed from the Comments collection by passing the comment's index.
     * After accessing the Comment object, customize the comment note by using the Comment object's setNote method.
     */
    public void addComment() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Adding a comment to "F5" cell
            int commentIndex = worksheet.getComments().add("F5");
            Comment comment = worksheet.getComments().get(commentIndex);

            //Setting the comment note
            comment.setNote("Hello Aspose!");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddComment_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add Comment", e);
        }
    }

    /**
     * It is possible to format the appearance of comments by configuring their height, width and font settings etc.
     */
    public void formatComment() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Adding a new worksheet to the Workbook object
            int sheetIndex = workbook.getWorksheets().add();
            Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

            //Adding a comment to "F5" cell
            int commentIndex = worksheet.getComments().add("F5");
            Comment comment = worksheet.getComments().get(commentIndex);

            //Setting the comment note
            comment.setNote("Hello Aspose!");

            //Setting the font size of a comment to 14
            Font font = comment.getFont();
            font.setSize(14);
            //Setting the font of a comment to bold
            font.setBold(true);
            //Setting the height of the font to 10
            comment.setHeightCM(10);
            //Setting the width of the font to 2
            comment.setWidthCM(2);

            //Saving the Excel file
            workbook.save(filePath + File.separator + "FormatComment_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Format Comment", e);
        }
    }
}
