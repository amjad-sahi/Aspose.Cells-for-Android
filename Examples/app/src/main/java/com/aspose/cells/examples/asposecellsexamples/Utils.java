package com.aspose.cells.examples.asposecellsexamples;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

import static android.content.ContentValues.TAG;

public class Utils {

    public static void saveFileToExternalStorage(Context context, String fileName) {

        try {
            AssetManager assetManager = context.getAssets();
            InputStream in = assetManager.open(fileName);

            FileOutputStream out = null;
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            myDir.mkdirs();

            File file = new File(myDir, fileName);
            file.createNewFile();
            if (file.exists()) file.delete();
            try {
                out = new FileOutputStream(file);
                int c;
                while ((c = in.read()) != -1) {
                    out.write(c);
                }
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (in != null) {
                    in.close();
                }
                if (out != null) {
                    out.close();
                }
            }
        } catch (Exception e) {
            Log.e(TAG, "Saving File To External Storage", e);
        }
    }
}
