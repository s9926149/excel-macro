package co.cynerds;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URISyntaxException;

public class ExcelMacro {
    public static void main(String[] args) throws URISyntaxException, IOException {
        String target = askPathsAndRun();
        System.out.println();
        System.out.printf("Successfully copied and pasted to <%s> %n", target);
//        new CopyPasteBot("./sample").pasteTo("./sample/target.xls");
    }

    public static String askPathsAndRun() throws IOException {
        BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

        System.out.print("Enter source files directory: ");
        String sourceDir = br.readLine().trim();

        System.out.println();
        System.out.print("Enter target template file path: ");
        String targetTemplateFilePath = br.readLine().trim();

        System.out.println();

        System.out.printf("Source: %s %n", sourceDir);
        System.out.printf("Target: %s %n", targetTemplateFilePath);

        System.out.println();

        System.out.print("Please confirm these are correct, Y/N? ");
        String confirm = br.readLine().trim();

        if (!"Y".equalsIgnoreCase(confirm)) {
            System.exit(1);
        }

        return new CopyPasteBot(sourceDir).pasteTo(targetTemplateFilePath);
    }
}