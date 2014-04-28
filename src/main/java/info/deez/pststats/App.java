package info.deez.pststats;

import java.io.File;
import java.io.FilenameFilter;
import java.io.PrintStream;
import java.util.Arrays;
import java.util.ArrayList;
import java.util.List;
import com.google.common.base.Preconditions;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;

public class App 
{
    public static void main(String[] args)
    {
        try {
            App app = new App();
            app.run(args);
        } catch (Exception e) {
            e.printStackTrace();
            System.exit(1);
        }
    }

    private Options options;

    public App() {
        this.options = new Options();
        options.addOption("d", "directory", true, "process PST files in specified directory");
        options.addOption("f", "file", true, "process specified PST file");
        options.addOption("h", "help", false, "print help and exit");
        options.addOption("o", "output", true, "write output to specified file");
    }

    public void run(String[] args) throws Exception
    {
        CommandLineParser parser = new BasicParser();
        CommandLine cmd = parser.parse(options, args);

        if (cmd.hasOption("h")) {
            printHelp();
            System.exit(0);
        }

        List<File> pstFiles = new ArrayList<File>();
        if (cmd.hasOption("d")) {
            File dir = new File(cmd.getOptionValue("d"));
            File[] files = dir.listFiles(new FilenameFilter() {
                public boolean accept(File dir, String name) {
                    return name.toLowerCase().endsWith(".pst");
                }
            });
            pstFiles.addAll(Arrays.asList(files));
        }
        if (cmd.hasOption("f")) {
            pstFiles.add(new File(cmd.getOptionValue("f")));
        }

        if (pstFiles.isEmpty()) {
            System.err.println("Error: One of -d or -f must be specified");
            printHelp();
            System.exit(1);
        }

        PrintStream os = System.out;
        if (cmd.hasOption("o")) {
            os = new PrintStream(cmd.getOptionValue("o"));
        }

        os.println("File,Num Recv Emails,Num Sent Emails");
        for (File pstFile : pstFiles) {
            processFile(pstFile, os);
        }
    }

    private void printHelp() {
        HelpFormatter formatter = new HelpFormatter();
        formatter.printHelp("pststats", options);
    }
    
    private void processFile(File file, PrintStream os) throws Exception {
        Preconditions.checkNotNull(file);

        PSTFile pstFile = new PSTFile(file);
        PSTFolder recvMailFolder = PSTFileUtils.findFolderByName(pstFile, "Received Email");
        if (recvMailFolder == null) {
            throw new RuntimeException("Could not find Received Email folder in PST file \"" + file + "\"");
        }
        PSTFolder sentMailFolder = PSTFileUtils.findFolderByName(pstFile, "Sent Items");
        if (sentMailFolder == null) {
            throw new RuntimeException("Could not find Sent Items folder in PST file \"" + file + "\"");
        }
        
        os.println(file + "," + recvMailFolder.getContentCount() + "," + sentMailFolder.getContentCount());
    }
}
