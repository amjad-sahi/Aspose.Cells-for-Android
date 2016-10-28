package com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadFormat;

import java.io.File;
import java.io.FileInputStream;

public class DetectFileFormat {

    private static final String TAG = DetectFileFormat.class.getName();

    public void detectFileFormat() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Detect the file format using file path
            int fmt = CellsHelper.detectFileFormat(filePath + File.separator + "source.xlsx");
            Log.v(TAG, "File Format Type is Xlsx: " + (fmt == FileFormatType.XLSX));

            FileInputStream fs = new FileInputStream(filePath);

            //Detect the file format using stream (i.e file stream, memory stream etc)
            fmt = CellsHelper.detectFileFormat(fs);
            Log.v(TAG, "File Format Type is Xlsx: " + (fmt == FileFormatType.XLSX));

            //Detect the load format using file path
            int ldfmt = CellsHelper.detectLoadFormat(filePath);
            Log.v(TAG, "Load Format is Xlsx: " + (ldfmt == LoadFormat.XLSX));

            //Close the stream
            fs.close();

            //Detect the file format using stream (i.e file stream, memory stream etc)
            fs = new FileInputStream(filePath);
            ldfmt = CellsHelper.detectLoadFormat(fs);
            Log.v(TAG, "Load Format is Xlsx: " + (ldfmt == LoadFormat.XLSX));

            //Close the stream
            fs.close();
        } catch (Exception e) {
            Log.e(TAG, "Detect File Format", e);
        }
    }
}
