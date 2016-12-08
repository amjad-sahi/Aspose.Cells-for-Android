package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Workbook;

import java.io.File;

public class CustomXMLParts {

    private static final String TAG = CustomXMLParts.class.getName();

    public void usingCustomXMLParts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            String booksXML= "<catalog>" +
                                "<book>" +
                                    "<title>Complete C#</title>" +
                                    "<price>44</price>" +
                                "</book>" +
                                "<book>" +
                                    "<title>Complete Java</title>" +
                                    "<price>76</price>" +
                                "</book>" +
                                "<book>" +
                                    "<title>Complete SharePoint</title>" +
                                    "<price>55</price>" +
                                "</book>" +
                                "<book>" +
                                    "<title>Complete PHP</title>" +
                                    "<price>63</price>" +
                                "</book>" +
                                "<book>" +
                                    "<title>Complete VB.NET</title>" +
                                    "<price>72</price>" +
                                "</book>" +
                             "</catalog>";

            Workbook workbook = new Workbook();
            workbook.getContentTypeProperties().add("BookStore", booksXML);
            workbook.save(filePath + "CustomXMLParts_Out.xlsx");

        } catch (Exception e) {
            Log.e(TAG, "Using Custom XML Parts in Aspose.Cells", e);
        }
    }
}
