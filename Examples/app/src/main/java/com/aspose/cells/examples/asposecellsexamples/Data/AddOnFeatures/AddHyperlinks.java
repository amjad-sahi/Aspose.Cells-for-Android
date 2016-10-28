package com.aspose.cells.examples.asposecellsexamples.Data.AddOnFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.FontUnderlineType;
import com.aspose.cells.HyperlinkCollection;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.File;

public class AddHyperlinks {

    private static final String TAG = AddHyperlinks.class.getName();

    public void addAURLLink() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet.
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet = worksheets.get(0);
            HyperlinkCollection hyperlinks = sheet.getHyperlinks();

            //Adding a hyperlink to a URL at "A1" cell
            hyperlinks.add("A1", 1, 1, "http://www.aspose.com");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddAURLLink_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a URL Link", e);
        }
    }

    public void applyFormattingToLookLikeHyperlink() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet.
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet = worksheets.get(0);

            //Setting a value to the "A1" cell
            Cells cells = sheet.getCells();
            Cell cell = cells.get("A1");
            cell.setValue("Visit Aspose");

            //Setting the font color of the cell to Blue
            Style style = cell.getStyle();
            style.getFont().setColor(Color.getBlue());

            //Setting the font of the cell to Single Underline
            style.getFont().setUnderline(FontUnderlineType.SINGLE);
            cell.setStyle(style);

            HyperlinkCollection hyperlinks = sheet.getHyperlinks();

            //Adding a hyperlink to a URL at "A1" cell
            hyperlinks.add("A1", 1, 1, "http://www.aspose.com");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "ApplyFormattingLikeHyperlink_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Apply Formatting to look like Hyperlink", e);
        }
    }

    public void addALinkToAnotherCellInTheSameFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet.
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet = worksheets.get(0);

            //Setting a value to the "A1" cell
            Cells cells = sheet.getCells();
            Cell cell = cells.get("A1");
            cell.setValue("Visit Aspose");

            //Setting the font color of the cell to Blue
            Style style = cell.getStyle();
            style.getFont().setColor(Color.getBlue());

            //Setting the font of the cell to Single Underline
            style.getFont().setUnderline(FontUnderlineType.SINGLE);
            cell.setStyle(style);

            HyperlinkCollection hyperlinks = sheet.getHyperlinks();

            //Adding an internal hyperlink to the "B9" cell of the other worksheet "Sheet2" in
            //the same Excel file
            hyperlinks.add("B3", 1, 1, "Sheet2!B9");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddALinkToAnotherCellInTheSameFile_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a Link to Another Cell in the Same File", e);
        }
    }

    public void addALinkToAnExternalFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the first worksheet.
            WorksheetCollection worksheets = workbook.getWorksheets();
            Worksheet sheet = worksheets.get(0);

            //Setting a value to the "A1" cell
            Cells cells = sheet.getCells();
            Cell cell = cells.get("A1");
            cell.setValue("Visit Aspose");

            //Setting the font color of the cell to Blue
            Style style = cell.getStyle();
            style.getFont().setColor(Color.getBlue());

            //Setting the font of the cell to Single Underline
            style.getFont().setUnderline(FontUnderlineType.SINGLE);
            cell.setStyle(style);

            HyperlinkCollection hyperlinks = sheet.getHyperlinks();

            //Adding a link to the external file
            hyperlinks.add("A5", 1, 1, filePath + File.separator + "Book1.xls");

            //Saving the Excel file
            workbook.save(filePath + File.separator + "AddALinkToAnExternalFile_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Add a Link to an External File", e);
        }
    }


}
