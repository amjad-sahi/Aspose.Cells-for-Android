package com.aspose.cells.examples.asposecellsexamples.Files;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

import java.io.File;
import java.io.InputStream;

public class OpenFile {

    private static final String TAG = OpenFile.class.getName();

    public void openThroughPath() {

        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an Workbook object with an Excel file path
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");
        } catch (Exception e) {
            Log.e(TAG, "Open through Path", e);
        }
    }

    public void openThroughStream(Context context) {
        try {
            AssetManager assetManager = context.getAssets();
            InputStream inputStream = assetManager.open("Book1.xlsx");

            //Creating an Workbook object with the stream object
            Workbook workbook = new Workbook(inputStream);
        } catch (Exception e) {
            Log.e(TAG, "Open through Stream", e);
        }
    }

    public void openMicrosoftExcel97File() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an EXCEL_97_TO_2003 LoadOptions object
            LoadOptions loadOptions = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

            //Creating an Workbook object with excel 97 file path and the loadOptions object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls", loadOptions);
        } catch (Exception e) {
            Log.e(TAG, "Open Microsoft Excel 97 Files", e);
        }
    }

    public void openMicrosoftExcel2007XLSXFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            LoadOptions loadOptions = new LoadOptions(FileFormatType.XLSX);

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx", loadOptions);
        } catch (Exception e) {
            Log.e(TAG, "Open Microsoft Excel 2007 XLSX Files", e);
        }
    }

    public void openCSVFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            LoadOptions loadOptions = new LoadOptions(FileFormatType.CSV);

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.csv", loadOptions);
        } catch (Exception e) {
            Log.e(TAG, "Open CSV Files", e);
        }
    }

    public void openTabDelimitedFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            LoadOptions loadOptions = new LoadOptions(FileFormatType.TAB_DELIMITED);

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.txt", loadOptions);
        } catch (Exception e) {
            Log.e(TAG, "Open Tab Delimited File", e);
        }
    }

    public void openEncryptedExcelFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Creating an EXCEL_97_TO_2003 LoadOptions object
            LoadOptions loadOptions = new LoadOptions(FileFormatType.EXCEL_97_TO_2003);

            //Setting the password for the encrypted Excel file
            loadOptions.setPassword("password1");

            //Creating an Workbook object with excel 97 file path and the loadOptions object
            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls", loadOptions);
        } catch (Exception e) {
            Log.e(TAG, "Open Encrypted Excel File", e);
        }
    }
}
