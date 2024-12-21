package com.kvm.Controller;

import com.kvm.Service.ExcelDataProcessorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

@Controller
@RequestMapping("/")
public class FileUploadController {

    @Autowired
    private ExcelDataProcessorService excelDataProcessorService;

    // This method is for the form page
    @GetMapping("/")
    public String uploadForm() {
        return "upload.html";
    }

    @PostMapping("/upload")
    public ResponseEntity<?> handleFileUpload(@RequestParam("file") MultipartFile file) {
        // Define a directory where uploaded files will be saved
        String uploadDir = System.getProperty("user.dir") + "/uploads/";

        // Create the directory if it doesn't exist
        File uploadDirectory = new File(uploadDir);
        if (!uploadDirectory.exists()) {
            uploadDirectory.mkdirs();  // Create the directory if it doesn't exist
        }

        // Define the file path for the uploaded file
        File inputFile = new File(uploadDir + "uploadedFile.xlsx");

        try {
            // Transfer the file to the specified path
            file.transferTo(inputFile);

            // Process the file (create a zip file)
            File zipFile = new File(uploadDir + "Filtered_Results.zip");
            excelDataProcessorService.processExcelData(inputFile, zipFile);

            // Ensure the file exists before sending it as a download
            if (zipFile.exists()) {
                byte[] fileContent = Files.readAllBytes(zipFile.toPath());

                return ResponseEntity.ok()
                        .header("Content-Disposition", "attachment; filename=" + zipFile.getName())
                        .contentType(MediaType.APPLICATION_OCTET_STREAM)
                        .body(fileContent);
            } else {
                return ResponseEntity.status(500).body("Error: Processed file not found");
            }
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body("Error processing the file: " + e.getMessage());
        }
    }
}
