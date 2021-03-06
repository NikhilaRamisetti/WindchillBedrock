package Windchill;

import com.relevantcodes.extentreports.ExtentReports;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.ZoneId;
import java.time.ZonedDateTime;

public class ReportFactory {

    private static File directory =  new File("");;
    private static String UserDirectory;

    static {
        try {
            UserDirectory = directory.getCanonicalPath();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static ExtentReports getInstance() throws IOException {
        ExtentReports extent;
      //  String Reportname = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());

        extent = new ExtentReports (UserDirectory +"/test-output/BedrockTestAutomationReport.html",false);
        extent
                .addSystemInfo("Host Name", "Windchill")
                .addSystemInfo("Environment", "Bedrock")
                .addSystemInfo("User Name", "User");
         extent.loadConfig(new File(UserDirectory+"/extent-config.xml"));
        return extent;
    }

    @Test
    public void DelExistingRep(){
        Path path
                = Paths.get(UserDirectory+"/test-output/BedrockTestAutomationReport.html");
        try {

            Files.deleteIfExists(path);
        }
        catch (Exception e) {
           // e.printStackTrace();
        }
    }
    @Test
    public void ReportProcess() throws IOException {
        try {
            {
                File file = new File(UserDirectory + "/test-output/BedrockTestAutomationReport.html");
                Path path = Paths.get(UserDirectory + "/test-output/BedrockTestAutomationReport.html");
                System.out.println(path);
                BasicFileAttributes attr;
                attr = Files.readAttributes(path, BasicFileAttributes.class);
                ZonedDateTime datetime=attr.creationTime().toInstant().atZone(ZoneId.systemDefault());
                String datetime1=datetime.toString();
                String replaceString=datetime1.replace(':','-');
                //System.out.println(replaceString);
                String filecreationdatetime=replaceString.substring(0,19);
                //System.out.println(filecreationdatetime);
                String lastEks = file.getName().toString();
                StringBuilder b = new StringBuilder(lastEks);
                String folderName=replaceString.substring(0,10);
                //System.out.println("lasteks :"+lastEks);
                //System.out.println("sb :"+b);
                String ReportArchivePath=UserDirectory + "/ReportArchive/" + folderName;
                File ReportArchivePath1 = new File(ReportArchivePath);
                if (!ReportArchivePath1.exists()) {
                   //System.out.print("No Folder");
                    ReportArchivePath1.mkdir();
                    //System.out.print("Folder created");
                }
                File backupDir = new File(ReportArchivePath);
                File temp = new File(backupDir.toString() + File.separator + file.getName().toString());
                if (file.getName().contains(".")) {
                    if (temp.exists()) {
                        temp = new File(backupDir.toString() + File.separator +
                                b.replace(lastEks.lastIndexOf("."), lastEks.lastIndexOf("."), filecreationdatetime).toString());
                    } else {
                        temp = new File(backupDir.toString() + File.separator +
                                b.replace(lastEks.lastIndexOf("."), lastEks.lastIndexOf("."), filecreationdatetime).toString());
                    }
                    b = new StringBuilder(temp.toString());
                } else {
                    temp = new File(backupDir.toString() + File.separator + file.getName());
                }

                if (temp.exists()) {
                    for (int x = 1; temp.exists(); x++) {
                        if (file.getName().contains(".")) {
                            temp = new File(b.replace(
                                    temp.toString().lastIndexOf(filecreationdatetime),
                                    temp.toString().lastIndexOf("."),
                                    filecreationdatetime + "(" +x+")").toString());
                        } else {
                            temp = new File(backupDir.toString() + File.separator
                                    + file.getName() + filecreationdatetime + "(" +x+")");
                        }
                    }
                    Files.copy(file.toPath(), temp.toPath());
                } else {
                    Files.copy(file.toPath(), temp.toPath());
                }
            }
            DelExistingRep();
        } catch (Exception e) {
            DelExistingRep();
            //e.printStackTrace();
        }
    }
    public ReportFactory() throws IOException {

    }
    }
