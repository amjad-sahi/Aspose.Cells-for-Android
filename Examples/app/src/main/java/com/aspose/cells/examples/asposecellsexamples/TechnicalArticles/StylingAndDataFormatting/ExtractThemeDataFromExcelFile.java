package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.Border;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.io.File;

public class ExtractThemeDataFromExcelFile {

    private static final String TAG = ExtractThemeDataFromExcelFile.class.getName();

    public void extractThemeDataFromExcelFile() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Create workbook object
            Workbook workbook = new Workbook(filePath + "source.xlsx");

            //Extract theme name applied to this workbook
            Log.i(TAG, workbook.getTheme());

            //Access first worksheet
            Worksheet worksheet = workbook.getWorksheets().get(0);

            //Access cell A1
            Cell cell = worksheet.getCells().get("A1");

            //Get the style object
            Style style = cell.getStyle();

            //Extract theme color applied to this cell
            Log.i(TAG, "Theme color: " + (style.getForegroundThemeColor().getColorType() == ThemeColorType.ACCENT_2));

            //Extract theme color applied to the bottom border of the cell
            Border bot = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
            Log.i(TAG, "Theme color: " + (bot.getThemeColor().getColorType() == ThemeColorType.ACCENT_1));

        } catch (Exception e) {
            Log.e(TAG, "Extract Theme Data from Excel File", e);
        }
    }
}
