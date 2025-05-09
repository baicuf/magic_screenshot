package com.tool;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.stream.Collectors;

public class ScreenshotToolApp {
    private static WebDriver driver;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(ScreenshotToolApp::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("Magic Screenshot Tool");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 300);
        frame.setLayout(new GridLayout(6, 1)); // +1 row for footer

        JButton selectCSV = new JButton("Select URLs.csv");
        JButton selectFolder = new JButton("Select Output Folder");
        JButton selectTemplate = new JButton("Select DOCX Template");
        JButton startButton = new JButton("Start");
        JLabel status = new JLabel("Waiting...", SwingConstants.CENTER);

        final Path[] csvPath = new Path[1];
        final Path[] outputDir = new Path[1];
        final Path[] docxTemplate = new Path[1];

        selectCSV.addActionListener((ActionEvent e) -> {
            JFileChooser fc = new JFileChooser();
            if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                csvPath[0] = fc.getSelectedFile().toPath();
                status.setText("CSV selected: " + csvPath[0].getFileName());
            }
        });

        selectFolder.addActionListener((ActionEvent e) -> {
            JFileChooser fc = new JFileChooser();
            fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                outputDir[0] = fc.getSelectedFile().toPath();
                status.setText("Output folder selected: " + outputDir[0].getFileName());
            }
        });

        selectTemplate.addActionListener((ActionEvent e) -> {
            JFileChooser fc = new JFileChooser();
            if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
                docxTemplate[0] = fc.getSelectedFile().toPath();
                status.setText("Template selected: " + docxTemplate[0].getFileName());
            }
        });

        startButton.addActionListener((ActionEvent e) -> {
            if (csvPath[0] == null || outputDir[0] == null || docxTemplate[0] == null) {
                JOptionPane.showMessageDialog(frame, "Please select all required files.");
                return;
            }
            try {
                run(csvPath[0], outputDir[0], docxTemplate[0]);
                status.setText("Done!");
            } catch (Exception ex) {
                ex.printStackTrace();
                status.setText("Error: " + ex.getMessage());
            }
        });

        // Footer label
        JLabel footer = new JLabel("developed by Florin Baicu, May 2025", SwingConstants.RIGHT);
        footer.setFont(new Font("SansSerif", Font.PLAIN, 10));

        frame.add(selectCSV);
        frame.add(selectFolder);
        frame.add(selectTemplate);
        frame.add(startButton);
        frame.add(status);
        frame.add(footer);
        frame.setVisible(true);
    }

    private static void run(Path csv, Path output, Path template) throws Exception {
        List<String> urls = Files.readAllLines(csv).stream()
                .map(String::trim).filter(line -> !line.isEmpty()).collect(Collectors.toList());

        Path chromeDriver = Paths.get("drivers", "chromedriver.exe");
        if (!Files.exists(chromeDriver)) {
            throw new FileNotFoundException("chromedriver.exe not found in 'drivers/'");
        }
        System.setProperty("webdriver.chrome.driver", chromeDriver.toAbsolutePath().toString());

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--force-device-scale-factor=1");
        options.addArguments("--window-size=1920,3000");
        driver = new ChromeDriver(options);

        JOptionPane.showMessageDialog(null, "Log in if needed. Then press OK to begin capturing.");

        // Determine dated output folder with suffix
        String dateStr = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        Path datedOutput = output.resolve(dateStr);
        int suffix = 1;
        while (Files.exists(datedOutput)) {
            datedOutput = output.resolve(dateStr + "_" + suffix);
            suffix++;
        }
        Files.createDirectories(datedOutput);

        // Load template once
        XWPFDocument doc;
        try (InputStream docStream = Files.newInputStream(template)) {
            doc = new XWPFDocument(docStream);
        }

        for (int i = 0; i < urls.size(); i++) {
            String url = urls.get(i);
            String urlPlaceholder = "$URL" + (i + 1);
            String imgPlaceholder = "$IMG" + (i + 1);

            driver.get(url);
            Thread.sleep(2000);

            Screenshot sc = new AShot().shootingStrategy(
                    ru.yandex.qatools.ashot.shooting.ShootingStrategies.viewportPasting(500)
            ).takeScreenshot(driver);

            BufferedImage fullImage = sc.getImage();
            File imageFile = datedOutput.resolve("screenshot_" + (i + 1) + ".png").toFile();
            ImageIO.write(fullImage, "PNG", imageFile);

            for (XWPFParagraph p : doc.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        if (text.contains(urlPlaceholder)) {
                            run.setText(text.replace(urlPlaceholder, url), 0);
                        }
                        if (text.contains(imgPlaceholder)) {
                            run.setText("", 0);  // Clear placeholder

                            int width = fullImage.getWidth();
                            int height = fullImage.getHeight();
                            int maxDocHeightPx = 3000; // conservative limit (in pixels)

                            int startY = 0;
                            int sliceIndex = 1;

                            while (startY < height) {
                                int sliceHeight = Math.min(maxDocHeightPx, height - startY);
                                BufferedImage slice = fullImage.getSubimage(0, startY, width, sliceHeight);

                                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                                ImageIO.write(slice, "png", baos);
                                InputStream is = new ByteArrayInputStream(baos.toByteArray());

                                int targetWidthEMU = Units.toEMU(500);
                                int scaledHeightEMU = (int) ((double) sliceHeight / width * targetWidthEMU);

                                run.addPicture(is, Document.PICTURE_TYPE_PNG,
                                        "slice_" + sliceIndex + ".png", targetWidthEMU, scaledHeightEMU);

                                startY += sliceHeight;
                                sliceIndex++;
                            }
                        }
                    }
                }
            }
        }

        driver.quit();

        // Save the final DOCX
        Path reportPath = datedOutput.resolve("report.docx");
        try (FileOutputStream out = new FileOutputStream(reportPath.toFile())) {
            doc.write(out);
        }
    }
}
