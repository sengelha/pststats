package info.deez.pststats;

import java.util.*;
import com.pff.*;

public class App 
{
    public static void main(String[] args)
    {
        try {
            App app = new App();
            app.run(args);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void run(String[] args) throws Exception
    {
        ArrayList<String> pstFiles = new ArrayList<String>();
        pstFiles.add("D:\\Steve\\Outlook\\2003.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2004.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2005.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2006.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2007.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2008.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2009.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2010.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2011.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2012.pst");
        pstFiles.add("D:\\Steve\\Outlook\\2013.pst");
        
        System.out.println("File,Num Recv Emails,Num Sent Emails");
        for (String pstFile : pstFiles) {
            processFile(pstFile);
        }
    }
    
    private void processFile(String pstFileName) throws Exception {
        String recvMailFolderPath = "/Top of Personal Folders/Received Email";        
        String sentMailFolderPath = "/Top of Personal Folders/Sent Items";        

        //System.out.println("Processing file " + pstFileName);
        PSTFile pstFile = new PSTFile(pstFileName);
        
        /*for (String folderName : enumAllFolderNames(pstFile)) {
            System.out.println(folderName);
        }*/
        
        PSTFolder recvMailFolder = getFolderByPath(pstFile, recvMailFolderPath);
        if (recvMailFolder == null) {
            throw new RuntimeException("Could not find folder \"" + recvMailFolderPath + "\" in PST file \"" + pstFileName + "\"");
        }
        PSTFolder sentMailFolder = getFolderByPath(pstFile, sentMailFolderPath);
        if (sentMailFolder == null) {
            throw new RuntimeException("Could not find folder \"" + sentMailFolderPath + "\" in PST file \"" + pstFileName + "\"");
        }
        
        System.out.println(pstFileName + "," + recvMailFolder.getContentCount() + "," + sentMailFolder.getContentCount());
    }
    
    private ArrayList<String> enumAllFolderNames(PSTFile pstFile) throws Exception {
        return enumAllFolderNames("/", pstFile.getRootFolder());
    }

    private ArrayList<String> enumAllFolderNames(String baseName, PSTFolder folder) throws Exception {
        ArrayList<String> folderNames = new ArrayList<String>();

        for (PSTFolder childFolder : folder.getSubFolders()) {
            for (String folderName : enumAllFolderNames(baseName + childFolder.getDisplayName() + "/", childFolder)) {
                folderNames.add(folderName);
            }
            folderNames.add(baseName + childFolder.getDisplayName());
        }
        
        return folderNames;
    }

    private PSTFolder getFolderByPath(PSTFile pstFile, String folderName) throws Exception
    {
        if (folderName.startsWith("/")) {
            return getFolderByPath(pstFile.getRootFolder(), folderName.substring(1));
        } else {
            return null;
        }        
    }
    
    private PSTFolder getFolderByPath(PSTFolder folder, String folderName) throws Exception {
        int indx = folderName.indexOf('/');
        if (indx == -1) {
            return getChildFolderByDisplayName(folder, folderName);
        } else {
            String childName = folderName.substring(0, indx);
            PSTFolder child = getChildFolderByDisplayName(folder, childName);
            return getFolderByPath(child, folderName.substring(indx + 1));
        }        
    }
    
    private PSTFolder getChildFolderByDisplayName(PSTFolder folder, String folderName) throws Exception {
        Vector<PSTFolder> childFolders = folder.getSubFolders();
        for (PSTFolder childFolder : childFolders) {
            if (childFolder.getDisplayName().equals(folderName))
                return childFolder;
        }
        return null;
    }
}
