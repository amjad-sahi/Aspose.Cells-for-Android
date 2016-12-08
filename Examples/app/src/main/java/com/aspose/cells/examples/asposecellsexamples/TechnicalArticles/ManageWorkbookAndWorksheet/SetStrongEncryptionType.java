package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.ManageWorkbookAndWorksheet;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;

import java.io.File;

public class SetStrongEncryptionType {

    private static final String TAG = SetStrongEncryptionType.class.getName();

    public void setStrongEncryptionType() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate a Workbook object.
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Password protect the file.
            workbook.getSettings().setPassword("1234");

            //Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
            workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

            //Save the Excel file.
            workbook.save(filePath + "EncryptedBook_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Set Strong Encryption Type", e);
        }
    }


}
