package com.aspose.cells.examples.asposecellsexamples.CellsHelperMethods;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;

import java.io.File;

public class MergeFiles {

    private static final String TAG = MergeFiles.class.getName();

    public void mergeFiles() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create an Array (length=2)
            String[] files = new String[2];
            //Specify files with their paths to be merged
            files[0] = filePath + File.separator + "Book1.xlsx";
            files[1] = filePath + File.separator + "source.xlsx";

            //Create a cachedFile for the process
            String cacheFile = filePath + File.separator + "temporaystoragefile.txt";
            //Output File to be created
            String dest = filePath + File.separator + "GrandBook_Out.xlsx";

            //Merge the files in the output file
            CellsHelper.mergeFiles(files, cacheFile, dest);

            //Now if you need to rename your sheets, you may load the output file
            Workbook workbook = new Workbook(filePath + File.separator + "GrandBook_Out.xlsx");

            int cnt = 1;
            //Browse all the sheets to rename them accordingly
            for( int i=0; i< workbook.getWorksheets().getCount();i++) {
                workbook.getWorksheets().get(i).setName("My_Custom_Sheet_" + cnt);
                cnt++;
            }

            //Re-save the file
            workbook.save(filePath + File.separator + "GrandBook_Out.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Merge Files", e);
        }
    }
}
