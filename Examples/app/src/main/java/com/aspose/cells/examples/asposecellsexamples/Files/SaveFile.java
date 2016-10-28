package com.aspose.cells.examples.asposecellsexamples.Files;

import android.content.Context;
import android.content.res.AssetManager;
import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

public class SaveFile {

    private static final String TAG = SaveFile.class.getName();


    public void saveToALocation() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xlsx");

            //Save in default (Excel2003) format
            workbook.save(filePath + File.separator + "book_out.xls");

            //Save in Excel2003 format
            workbook.save(filePath + File.separator + "book_out.xls", FileFormatType.EXCEL_97_TO_2003);

            //Save in Excel2007 xlsx format
            workbook.save(filePath + File.separator + "book_out.xlsx", FileFormatType.XLSX);

            //Save in SpreadsheetML format
            workbook.save(filePath + File.separator + "book_out.xml", FileFormatType.EXCEL_2003_XML);

        } catch (Exception e) {
            Log.e(TAG, "Save to a location", e);
        }
    }

    public void saveToAStream(Context context) {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Create a Stream object
            FileOutputStream stream = new FileOutputStream(filePath + File.separator + "book_out.xls");

            AssetManager assetManager = context.getAssets();
            InputStream inputStream = assetManager.open("Book1.xlsx");
            //Create a Workbook object with the stream object
            Workbook workbook = new Workbook(inputStream);

            //Save in Excel2003 format
            workbook.save(stream, FileFormatType.EXCEL_97_TO_2003);

            //Save in MS Excel2007 xlsx format
            workbook.save(stream, FileFormatType.XLSX);

            //Save in SpreadsheetML format
            workbook.save(stream, FileFormatType.EXCEL_2003_XML);
        } catch (Exception e) {
            Log.e(TAG, "Save to a Stream", e);
        }
    }

    public void saveEntireWorkbookIntoTextOrCSVFormat() {

        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            //Load source workbook
            Workbook workbook = new Workbook(filePath + File.separator + "source.xlsx");

            //0-byte array
            byte[] workbookData = new byte[0];

            //Text save options. You can use any type of separator
            TxtSaveOptions opts = new TxtSaveOptions();
            opts.setSeparator('\t');

            //Copy each worksheet data in text format inside workbook data array
            for (int idx = 0; idx < workbook.getWorksheets().getCount(); idx++) {
                //Save the active worksheet into text format
                ByteArrayOutputStream bout = new ByteArrayOutputStream();
                workbook.getWorksheets().setActiveSheetIndex(idx);
                workbook.save(bout, opts);

                //Save the worksheet data into sheet data array
                byte[] sheetData = bout.toByteArray();

                //Combine this worksheet data into workbook data array
                byte[] combinedArray = new byte[workbookData.length + sheetData.length];
                System.arraycopy(workbookData, 0, combinedArray, 0, workbookData.length);
                System.arraycopy(sheetData, 0, combinedArray, workbookData.length, sheetData.length);

                workbookData = combinedArray;
            }

            //Save entire workbook data into file
            FileOutputStream fout = new FileOutputStream(filePath + File.separator + "EntireWorkbookIntoText_Out.txt");
            fout.write(workbookData);
            fout.close();

        } catch (Exception e) {
            Log.e(TAG, "Save entire Workbook into Text or CSV Format", e);
        }
    }
}
