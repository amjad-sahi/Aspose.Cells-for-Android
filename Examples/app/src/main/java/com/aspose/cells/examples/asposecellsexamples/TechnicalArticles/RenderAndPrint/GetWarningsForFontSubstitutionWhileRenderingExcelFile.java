package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.IWarningCallback;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.WarningInfo;
import com.aspose.cells.WarningType;
import com.aspose.cells.Workbook;

import java.io.File;

public class GetWarningsForFontSubstitutionWhileRenderingExcelFile {

    private static final String TAG = GetWarningsForFontSubstitutionWhileRenderingExcelFile.class.getName();

    public void getWarningsForFontSubstitution() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            Workbook workbook = new Workbook(filePath + "source.xlsx");
            PdfSaveOptions options = new PdfSaveOptions();
            options.setWarningCallback(new WarningCallback());
            workbook.save(filePath + "WarningsForFontSubstitution_Out.pdf", options);
        } catch (Exception e) {
            Log.e(TAG, "Get Warnings for Font Substitution while Rendering Excel File", e);
        }
    }

    public class WarningCallback implements IWarningCallback {

        @Override
        public void warning(WarningInfo info) {
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION) {
                Log.e(TAG, "WARNING INFO: " + info.getDescription());
            }
        }
    }
}
