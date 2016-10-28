package com.aspose.cells.examples.asposecellsexamples.UtilityFeatures;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.EncryptionType;
import com.aspose.cells.Workbook;

import java.io.File;

public class EncryptFile {

    private static final String TAG = EncryptFile.class.getName();

    public void encryptFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath();

            Workbook workbook = new Workbook(filePath + File.separator + "Book1.xls");

            //Password protect the file.
            workbook.getSettings().setPassword("1234");

            //Specify XOR encryption type.
            workbook.setEncryptionOptions(EncryptionType.XOR, 40);

            //Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
            workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);

            //Save the Excel file.
            workbook.save(filePath + File.separator + "EncryptedBook1_out.xls");
        } catch(Exception e) {
            Log.e(TAG, "Encrypt File", e);
        }
    }
}
