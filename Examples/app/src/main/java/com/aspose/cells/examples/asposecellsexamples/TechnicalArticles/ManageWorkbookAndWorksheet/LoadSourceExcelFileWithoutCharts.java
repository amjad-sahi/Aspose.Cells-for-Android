package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;

public class LoadSourceExcelFileWithoutCharts {

    private static final String TAG = LoadSourceExcelFileWithoutCharts.class.getName();

    public void loadSourceExcelFileWithoutCharts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Specify the load options and filter the data
            LoadOptions options = new LoadOptions();
            options.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.CHART);

            //Load the workbook with specified load options
            Workbook workbook = new Workbook(filePath + "sample.xlsx", options);

            //Save the workbook in output format
            workbook.save(filePath + "loadFileWithoutCharts_Out.pdf", SaveFormat.PDF);
        } catch (Exception e) {
            Log.e(TAG, "Load Source Excel File Without Charts", e);
        }
    }
}
