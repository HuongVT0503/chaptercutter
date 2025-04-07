package com.example.chaptercutter;

//import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;


import java.util.List;

public class HelloController {
    public TextField limiting;
    public Label statusLabel;
    @FXML




    // Handler for the Choose File button

    protected void handleChooseFile() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Word Documents", "*.docx"));
        File selectedFile = fileChooser.showOpenDialog(new Stage());
        if (selectedFile != null) {
            // Process the file and update status
            File outputFile = new File(selectedFile.getParent(), "processed_" + selectedFile.getName());
            boolean success = processDocx(selectedFile, outputFile);
            if (success) {
                statusLabel.setText("File processed successfully: " + outputFile.getName());
            } else {
                statusLabel.setText("Error processing file.");
            }
        } else {
            statusLabel.setText("No file selected.");
        }
    }

    // Method to process the .docx file based on your requirements
    private boolean processDocx(File inputFile, File outputFile) {



        // int wordLimit = Integer.parseInt(limiting.getText());
        final int WORD_LIMIT =  Integer.parseInt(limiting.getText());
        ZipSecureFile.setMinInflateRatio(0.001);
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis)) {

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            StringBuilder textBuffer = new StringBuilder();
            int wordCount = 0;
            int chapterNumber = 1;
            int startIndex = 0;

            // Check if the first paragraph is only digits
            if (!paragraphs.isEmpty()) {
                String firstLine = paragraphs.getFirst().getText().trim();
                if (firstLine.matches("^\\d+$")) { // first paragraph is an int
                    chapterNumber = Integer.parseInt(firstLine);
                    // Output the first paragraph (the number) as is
                    textBuffer.append(firstLine).append("\n\n");
                    // Immediately insert a chapter marker using that number
                    textBuffer.append("Chương ").append(chapterNumber).append("\n\n");
                    chapterNumber++; // Next chapter will be chapterNumber+1
                    startIndex = 1; // Skip the first paragraph since it’s already processed
                }
            }

            // Process subsequent paragraphs
            for (int i = startIndex; i < paragraphs.size(); i++) {
                XWPFParagraph para = paragraphs.get(i);
                String text = para.getText();
                System.out.println("Paragraph: " + text);

                // Count valid tokens (tokens containing at least one letter or digit)
                String[] tokens = text.split("\\s+");
                int countTokens = 0;
                for (String token : tokens) {
                    if (token.matches(".*[A-Za-z0-9].*")) {
                        countTokens++;
                    }
                }
                wordCount += countTokens;
                textBuffer.append(text).append("\n");

                // If the paragraph ended and the accumulated word count meets or exceeds the limit, insert a chapter marker.
                if (wordCount >= WORD_LIMIT) {
                    textBuffer.append("\n\nChương ").append(chapterNumber).append("\n\n");
                    chapterNumber++;
                    wordCount = 0; // reset word count after inserting the marker
                }
            }

            // Write the output into a new document, preserving paragraph breaks.
            try (XWPFDocument newDoc = new XWPFDocument();
                 FileOutputStream fos = new FileOutputStream(outputFile)) {

                String[] lines = textBuffer.toString().split("\n");
                for (String line : lines) {
                    newDoc.createParagraph().createRun().setText(line);
                }
                newDoc.write(fos);
            }
            return true;
        } catch (IOException e) {
            // e.printStackTrace();
            return false;
        }
    }
}
