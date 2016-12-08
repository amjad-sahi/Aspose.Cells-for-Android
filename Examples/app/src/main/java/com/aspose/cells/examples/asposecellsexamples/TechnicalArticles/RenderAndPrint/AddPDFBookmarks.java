package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Cell;
import com.aspose.cells.PdfBookmarkEntry;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;

import java.io.File;
import java.util.ArrayList;

public class AddPDFBookmarks {

    private static final String TAG = AddPDFBookmarks.class.getName();

    public void addPDFBookmarks() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a new workbook
            Workbook workbook = new Workbook();

            //Get the worksheets in the workbook
            WorksheetCollection worksheets = workbook.getWorksheets();

            //Add a sheet to the workbook
            worksheets.add("1");

            //Add 2nd sheet to the workbook
            worksheets.add("2");

            //Add the third sheet
            worksheets.add("3");

            //Get cells in different worksheets
            Cell cellInPage1 = worksheets.get(0).getCells().get("A1");
            Cell cellInPage2 = worksheets.get(1).getCells().get("A1");
            Cell cellInPage3 = worksheets.get(2).getCells().get("A1");

            //Add a value to the A1 cell in the first sheet
            cellInPage1.setValue("a");

            //Add a value to the A1 cell in the second sheet
            cellInPage2.setValue("b");

            //Add a value to the A1 cell in the third sheet
            cellInPage3.setValue("c");

            //Create the PdfBookmark entry object
            PdfBookmarkEntry pbeRoot = new PdfBookmarkEntry();

            //Set its text
            pbeRoot.setText("root");

            //Set its destination source page
            pbeRoot.setDestination(cellInPage1);

            //Set the bookmark collapsed
            pbeRoot.setOpen(false);

            //Add a new PdfBookmark entry object
            PdfBookmarkEntry subPbe1 = new PdfBookmarkEntry();

            //Set its text
            subPbe1.setText("1");

            //Set its destination source page
            subPbe1.setDestination(cellInPage2);

            //Add another PdfBookmark entry object
            PdfBookmarkEntry subPbe2 = new PdfBookmarkEntry();

            //Set its text
            subPbe2.setText("2");

            //Set its destination source page
            subPbe2.setDestination(cellInPage3);

            //Create an array list
            ArrayList subEntryList = new ArrayList();

            //Add the entry objects to it
            subEntryList.add(subPbe1);
            subEntryList.add(subPbe2);
            pbeRoot.setSubEntry(subEntryList);

            //Set the PDF bookmarks
            PdfSaveOptions options = new PdfSaveOptions();
            options.setBookmark(pbeRoot);

            //Save the PDF file
            workbook.save(filePath + "AddPDFBookmarks_Out.pdf", options);
        } catch (Exception e) {
            Log.e(TAG, "Adding PDF Bookmarks", e);
        }
    }
}
