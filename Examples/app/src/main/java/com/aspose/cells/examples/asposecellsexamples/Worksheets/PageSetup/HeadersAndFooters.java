package com.aspose.cells.examples.asposecellsexamples.Worksheets.PageSetup;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;

import java.io.File;
import java.io.FileInputStream;

public class HeadersAndFooters {
    private static final String TAG = HeadersAndFooters.class.getName();

    public void setHeadersAndFooters() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Instantiating a Workbook object
            Workbook workbook = new Workbook();

            //Obtaining the reference of the PageSetup of the worksheet
            PageSetup pageSetup = workbook.getWorksheets().get(0).getPageSetup();

            //Setting worksheet name at the left  header
            pageSetup.setHeader(0, "&A");

            //Setting current date and current time at the central header
            //and changing the font of the header
            pageSetup.setHeader(1, "&\"Times New Roman,Bold\"&D-&T");

            //Setting current file name at the right header and changing the font of the header
            pageSetup.setHeader(2, "&\"Times New Roman,Bold\"&12&F");

            //Setting a string at the left footer and changing the font of the footer
            pageSetup.setFooter(0, "Hello World! &\"Courier New\"&14 123");

            //Setting picture at the central footer
            pageSetup.setFooter(1, "&G");

            FileInputStream fis = new FileInputStream(filePath + File.separator + "footer.jpg");
            byte[] picData = new byte[fis.available()];
            fis.read(picData);
            pageSetup.setFooterPicture(1, picData);
            fis.close();

            //Setting the current page number and page count at the right footer
            pageSetup.setFooter(2, "&Pof&N");

            workbook.save(filePath + File.separator + "SetHeadersAndFooters_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Freeze Panes", e);
        }
    }

    public void insertAGraphicInAHeaderOrFooter() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating a Workbook object
            Workbook workbook = new Workbook();

            //Creating a string variable to store the URL of the logo/picture
            String logo_url = filePath + File.separator + "school.jpg";

            //Creating the instance of the FileInputStream object to open the logo/picture in the stream
            FileInputStream inFile = new FileInputStream(logo_url);

            //Creating a PageSetup object to get the page settings of the first worksheet of the workbook
            PageSetup pageSetup = workbook.getWorksheets().get(0).getPageSetup();

            //Setting the logo or picture in the central section of the page header
            pageSetup.setHeader(1, "&G");
            byte[] picData = new byte[inFile.available()];
            inFile.read(picData);
            pageSetup.setHeaderPicture(1,picData);

            //Setting the Sheet's name in the right section of the page header with the script
            pageSetup.setHeader(2, "&A");

            //Saving the workbook
            workbook.save(filePath + File.separator + "InsertGraphicInAHeaderOrFooter_Out.xls");

            //Closing the FileStream object
            inFile.close();
        } catch (Exception e) {
            Log.e(TAG, "Insert a Graphic in a Header or Footer", e);
        }
    }

}
