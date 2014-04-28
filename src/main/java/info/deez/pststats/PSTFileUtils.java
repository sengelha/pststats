package info.deez.pststats;

import com.pff.PSTFile;
import com.pff.PSTFolder;

public class PSTFileUtils {
    public static PSTFolder findFolderByName(PSTFile file, String displayName) throws Exception {
        return findFolderByName(file.getRootFolder(), displayName);
    }

    private static PSTFolder findFolderByName(PSTFolder baseFolder, String displayName) throws Exception {
        if (baseFolder.getDisplayName().equals(displayName))
            return baseFolder;

        for (PSTFolder childFolder : baseFolder.getSubFolders())
        {
            PSTFolder foundFolder = findFolderByName(childFolder, displayName);
            if (foundFolder != null)
                return foundFolder;
        }

        return null;
    }
}