package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.IFilePathProvider;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.asposecellsexamples.Utils;

import java.io.File;

public class ProvideExportedWorksheetHtmlFilePath {

    private static final String TAG = ProvideExportedWorksheetHtmlFilePath.class.getName();

    public void provideExportedWorksheetHtmlFilePath() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //If you will not set the license, program will go in infinite loop
            //because Aspose.Cells will always make the warning worksheet as active
            //sheet in Evaluation mode.
            Utils.applyALicense();

            //Check if license is set, otherwise do not proceed
            Workbook wb = new Workbook();

            if (wb.isLicensed() == false) {
                Log.e(TAG, "You must set the license to execute this code successfully.");
            } else {
                //Test IFilePathProvider interface
                TestIFilePathProvider pg = new TestIFilePathProvider();
                pg.TestFilePathProvider(filePath);
                Log.i(TAG, "Done.");
            }
        } catch (Exception e) {
            Log.e(TAG, "Provide exported worksheet html file path via IFilePathProvider interface", e);
        }
    }

    public class TestIFilePathProvider {

        String mDirPath;

        //Implementation of IFilePathProvider interface
        public class FilePathProvider implements IFilePathProvider {
            //Gets the full path of the file by worksheet name when exporting worksheet to html separately.
            //So the references among the worksheets could be exported correctly.
            public String getFullName(String sheetName) {

                if ("Sheet2".equals(sheetName)) {
                    return mDirPath + "OtherSheets" + File.separator + "Sheet2.html";
                } else if ("Sheet3".equals(sheetName)) {
                    return mDirPath + "OtherSheets" + File.separator + "Sheet3.html";
                }
                return "";
            }
        }

        void TestFilePathProvider(String filePath) throws Exception {

            mDirPath = filePath;
            //Create subdirectory for second and third worksheets
            File dir = new File(mDirPath + "OtherSheets");
            dir.mkdir();

            //Load sample workbook from your directory
            Workbook wb = new Workbook(mDirPath + "IFilePathProvider_Sample.xlsx");

            //Save worksheets to separate html files
            //Because of IFilePathProvider, hyperlinks will not be broken.
            for (int i = 0; i < wb.getWorksheets().getCount(); i++) {
                //Set the active worksheet to current value of variable i
                wb.getWorksheets().setActiveSheetIndex(i);

                //Create html save option
                HtmlSaveOptions options = new HtmlSaveOptions();
                options.setExportActiveWorksheetOnly(true);

                //If you will comment this line, then hyperlinks will be broken
                options.setFilePathProvider(new FilePathProvider());

                //Sheet actual index which starts from 1 not from 0
                int sheetIndex = i + 1;

                String completeFilePath = "";

                //Save first sheet to same directory and second and third worksheets to subdirectory
                if (i == 0) {
                    completeFilePath = mDirPath + "Sheet1.html";
                } else {
                    completeFilePath = mDirPath + "OtherSheets" + File.separator + "Sheet" + sheetIndex + ".html";
                }

                //Save the worksheet to html file
                wb.save(completeFilePath, options);
            }
        }
    }
}

