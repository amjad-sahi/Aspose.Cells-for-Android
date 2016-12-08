package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.StylingAndDataFormatting;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;

import java.io.File;

public class Excel2007ThemesAndColors {

    private static final String TAG = Excel2007ThemesAndColors.class.getName();

    public void getAndSetThemeColors() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate an instance of Workbook &
            //load an exiting spreadsheet
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Get the Background1 theme color
            Color color = workbook.getThemeColor(ThemeColorType.BACKGROUND_1);

            //Get the Accent2 theme color
            color = workbook.getThemeColor(ThemeColorType.ACCENT_1);

            //Change the Background1 theme color
            workbook.setThemeColor(ThemeColorType.BACKGROUND_1, Color.getRed());

            //Get the updated Background1 theme color
            color = workbook.getThemeColor(ThemeColorType.BACKGROUND_1);

            //Change the Accent2 theme color
            workbook.setThemeColor(ThemeColorType.ACCENT_1, Color.getBlue());

            //Get the updated Accent2 theme color
            color = workbook.getThemeColor(ThemeColorType.ACCENT_1);

            //Save the updated file
            workbook.save(filePath + "GetAndSetThemeColors_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Get and Set Theme Colors", e);
        }
    }

    public void applyCustomThemes() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Define Color array (of 12 colors) for the Theme
            Color[] carr = new Color[12];
            carr[0] = Color.getAntiqueWhite(); // Background1
            carr[1] = Color.getBrown(); // Text1
            carr[2] = Color.getAliceBlue(); // Background2
            carr[3] = Color.getYellow(); // Text2
            carr[4] = Color.getYellowGreen(); // Accent1
            carr[5] = Color.getRed(); // Accent2
            carr[6] = Color.getPink(); // Accent3
            carr[7] = Color.getPurple(); // Accent4
            carr[8] = Color.getPaleGreen(); // Accent5
            carr[9] = Color.getOrange(); // Accent6
            carr[10] = Color.getGreen(); // Hyperlink
            carr[11] = Color.getGray(); // Followed Hyperlink

            //Instantiate an instance of Workbook &
            //load a spreadsheet file
            Workbook workbook = new Workbook(filePath + "Book1.xls");

            //Set the custom theme with specified colors
            workbook.customTheme("CustomeTheme1", carr);

            //Save as the excel file
            workbook.save(filePath + "CustomThemes_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Apply custom themes", e);
        }
    }

    public void usingThemeColors() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Instantiate an instance of Workbook
            Workbook workbook = new Workbook();

            //Get cells collection in the first (default) worksheet
            Cells cells = workbook.getWorksheets().get(0).getCells();

            //Get the D3 cell
            Cell cell = cells.get("D3");

            //Get the style of the cell
            Style style = cell.getStyle();

            //Set background color for the cell from the default theme Accent2 color
            style.setBackgroundThemeColor(new ThemeColor(ThemeColorType.ACCENT_2, 0.5));

            //Set the pattern type
            style.setPattern(BackgroundType.SOLID);

            //Get the font for the style
            Font font = style.getFont();

            //Set the theme color
            font.setThemeColor(new ThemeColor(ThemeColorType.ACCENT_4, 0.1));

            //Apply style
            cell.setStyle(style);

            //Put a value
            cell.putValue("Testing");

            //Save the excel file
            workbook.save(filePath + "UsingThemeColors_Out.xls");
        } catch (Exception e) {
            Log.e(TAG, "Using Theme Colors", e);
        }
    }
}
