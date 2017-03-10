package console;

import gui.GUI;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Main {
    private static final String encoding = "UTF-8";

    public static void main(String[] args) throws IOException {
        System.out.println(getLastDir());
        GUI window = new GUI(getLastDir());
        SwingUtilities.invokeLater(window);
    }

    private static String getLastDir() throws IOException {
        byte[] encoded = Files.readAllBytes(Paths.get("last_dir.in"));
        String dir = new String(encoded, encoding);
        Path path = Paths.get(dir);
        return Files.exists(path)? dir : System.getProperty("user.home");
    }

    public static void setLastDir(String dir) {
        PrintWriter writer = null;
        try {
            writer = new PrintWriter("last_dir.in");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        writer.print(dir);
        writer.close();
    }
}
