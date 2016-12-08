package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatInfo;
import com.aspose.cells.FileFormatUtil;

import java.io.File;

public class DetectAFileFormatAndCheckIfFileIsEncrypted {

    private static final String TAG = DetectAFileFormatAndCheckIfFileIsEncrypted.class.getName();

    public void detectAFileFormatAndCheckIfTheFileIsEncrypted() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Detect file format
            FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath + "Book1.xlsx");

            //Gets the detected load format
            Log.i(TAG, "The spreadsheet format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));

            //Check if the file is encrypted.
            Log.i(TAG, "The file is encrypted: " + info.isEncrypted());
        } catch (Exception e) {
            Log.e(TAG, "Detect a File Format and Check if the File is Encrypted", e);
        }
    }

}
