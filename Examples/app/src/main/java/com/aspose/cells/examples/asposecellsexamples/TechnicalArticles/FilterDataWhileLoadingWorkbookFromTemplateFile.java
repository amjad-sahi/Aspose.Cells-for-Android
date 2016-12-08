package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class FilterDataWhileLoadingWorkbookFromTemplateFile {

    private static final String TAG = FilterDataWhileLoadingWorkbookFromTemplateFile.class.getName();

    public void filterDataWhileLoadingWorkbookFromTemplateFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Set the load options, we only want to load shapes and do not want to load data
            LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
            opts.setLoadDataFilterOptions(LoadDataFilterOptions.SHAPE);

            //Create workbook object from sample excel file using load options
            Workbook wb = new Workbook(filePath + "templateFile.xlsx", opts);

            //Save the output in pdf format
            wb.save(filePath + "FilterData_Out.pdf", SaveFormat.PDF);
        } catch (Exception e) {
            Log.e(TAG, "Filtering the kind of data while loading the workbook from template file", e);
        }

    }
}
