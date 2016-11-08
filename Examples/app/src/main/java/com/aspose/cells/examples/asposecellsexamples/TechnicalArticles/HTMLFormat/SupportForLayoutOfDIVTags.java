package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.annotation.TargetApi;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

public class SupportForLayoutOfDIVTags {

    private static final String TAG = SupportForLayoutOfDIVTags.class.getName();

    /**
     * Support for Layout of DIV Tags while Loading HTML.
     */
    @TargetApi(19)
    public void supportForLayoutOfDIVTagsWhileLoadingHTML() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Store the HTML snippet in a variable
            String export_html = "<html>"
                    + "<body>"
                    + " <table>"
                    + "     <tr>"
                    + "         <td>"
                    + "             <div>This is some Text.</div>"
                    + "             <div>"
                    + "                 <div>"
                    + "                     <span>This is some more Text</span>"
                    + "                 </div>"
                    + "                 <div>"
                    + "                     <span>abc@abc.com</span>"
                    + "                 </div>"
                    + "                 <div>"
                    + "                     <span>1234567890</span>"
                    + "                 </div>"
                    + "                 <div>"
                    + "                     <span>ABC DEF</span>"
                    + "                 </div>"
                    + "             </div>"
                    + "             <div>Generated On May 30, 2016 02:33 PM <br />Time Call Received from Jan 01, 2016 to May 30, 2016</div>"
                    + "         </td>"
                    + "         <td>"
                    + "             <img src='ASpose_logo_100x100.png' />"
                    + "         </td>"
                    + "     </tr>"
                    + " </table>"
                    + "</body>"
                    + "</html>";

            //Convert HTML string to InputStream
            InputStream stream = new ByteArrayInputStream(export_html.getBytes(StandardCharsets.UTF_8));

            //Create an instance of HTMLLoadOptions
            HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
            // Set SupportDivTag property to true
            loadOptions.setSupportDivTag(true);

            //Create an instance of Workbook from stream
            Workbook book = new Workbook(stream, loadOptions);
            //Save the spreadsheet in HTML format
            book.save(filePath + "SupportForLayout_Out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            Log.e(TAG, "Support for Layout of DIV Tags while Loading HTML", e);
        }
    }
}
