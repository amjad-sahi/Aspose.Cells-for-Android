package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.HTMLFormat;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;

import java.io.File;

public class ChangeTheHTMLLinkTargetType {

    private static final String TAG = ChangeTheHTMLLinkTargetType.class.getName();

    public void changeTheHTMLLinkTargetType() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "HTMLLink.xlsx");

            HtmlSaveOptions opts = new HtmlSaveOptions();
            opts.setLinkTargetType(HtmlLinkTargetType.SELF);

            workbook.save(filePath + "ChangeTheHTMLLinkTargetType_Out.html", opts);
        } catch (Exception e) {
            Log.e(TAG, "Change the HTML Link Target Type", e);
        }
    }
}
