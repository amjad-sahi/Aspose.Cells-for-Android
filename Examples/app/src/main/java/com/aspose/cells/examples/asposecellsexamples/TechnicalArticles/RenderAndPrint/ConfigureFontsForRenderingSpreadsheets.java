package com.aspose.cells.examples.asposecellsexamples.TechnicalArticles.RenderAndPrint;

import android.os.Environment;
import android.util.Log;

import com.aspose.cells.FileFontSource;
import com.aspose.cells.FolderFontSource;
import com.aspose.cells.FontConfigs;
import com.aspose.cells.FontSourceBase;
import com.aspose.cells.MemoryFontSource;

import java.io.File;
import java.io.FileInputStream;

public class ConfigureFontsForRenderingSpreadsheets {

    private static final String TAG = ConfigureFontsForRenderingSpreadsheets.class.getName();

    public void selectionOfFonts() {
        try {
            String root = Environment.getExternalStorageDirectory().toString();
            File myDir = new File(root + File.separator + "Aspose");
            String filePath = myDir.getCanonicalPath() + File.separator;

            //Defining string variables to store paths to font folders & font file
            String fontFolder1 = filePath + "Arial";
            String fontFolder2 = filePath + "Calibri";
            String fontFile = filePath + "Arial/arial.ttf";

            //Setting first font folder with setFontFolder method
            //Second parameter directs the API to search the sub folders for font files
            FontConfigs.setFontFolder(fontFolder1, true);

            //Setting both font folders with setFontFolders method
            //Second parameter prohibits the API to search the sub folders for font files
            FontConfigs.setFontFolders(new String[] { fontFolder1, fontFolder2 }, false);

            //Defining FolderFontSource
            FolderFontSource sourceFolder = new FolderFontSource(fontFolder1, false);

            //Defining FileFontSource
            FileFontSource sourceFile = new FileFontSource(fontFile);

            //Defining MemoryFontSource
            File file = new File(fontFile);
            FileInputStream fin = new FileInputStream(file);
            byte[] bytes = new byte[(int)file.length()];
            fin.read(bytes);
            fin.close();
            MemoryFontSource sourceMemory = new MemoryFontSource(bytes);

            //Setting font sources
            FontConfigs.setFontSources(new FontSourceBase[] { sourceFolder, sourceFile, sourceMemory });

        } catch (Exception e) {
            Log.e(TAG, "Selection of Fonts", e);
        }
    }

    /**
     * Aspose.Cells APIs also provide the ability to specify the substitute font for rendering purposes.
     *  This mechanism is helpful when a required font is not available on the machine where conversion has to take place.
     *  Users can provide a list of font names as alternative to the originally required font.
     *  In order to achieve this, the Aspose.Cells APIs have exposed the FontConfigs.setFontSubstitutes method which accepts 2 parameters.
     *  The first parameter is of type String, which should be the name of font which needs to be substituted.
     *  The second parameter is an array of type String.
     *  Users can provide a list of font names as substitution to original font (specified in the first parameter).
     */
    public void fontSubstitutionMechanism() {
        //Substituting the Arial font with Times New Roman & Calibri
        FontConfigs.setFontSubstitutes("Arial", new String[] { "Times New Roman", "Calibri" });
    }
}
