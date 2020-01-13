package com.example.xlsreader.controller;

import com.github.pjfanning.xlsx.StreamingReader;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiImplicitParams;
import org.apache.poi.ss.usermodel.*;
import org.apache.tomcat.util.http.fileupload.FileItemIterator;
import org.apache.tomcat.util.http.fileupload.FileItemStream;
import org.apache.tomcat.util.http.fileupload.FileUploadException;
import org.apache.tomcat.util.http.fileupload.servlet.ServletFileUpload;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

@RestController
public class democontroller {

    @RequestMapping("/hello/{name}")
    String home(@PathVariable String name) {


        return ( "Hello World! "+name);
    }

    @ApiImplicitParams(value = {@ApiImplicitParam(dataType = "file", name = "file", required = true,paramType = "form")})
    @PostMapping(value="/upload",consumes = {MediaType.MULTIPART_FORM_DATA_VALUE})
    public String upload(MultipartFile file) throws IOException, FileUploadException {
        File file1 = saveToFile(file).toFile();
        InputStream is = new FileInputStream(file1);
        Workbook workbook = StreamingReader.builder()
                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);
        for (Sheet sheet : workbook){
            System.out.println(sheet.getSheetName());
            for (Row r : sheet) {
                for (int y = 0; y < 4; y++) {
                    System.out.println(getvalue(r.getCell(y)));
                }
            }
        }
        return "Ok";
    }

    public Object getvalue(Cell cell) {
        Object obj =null;
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    obj = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if(DateUtil.isCellDateFormatted(cell))
                        obj=cell.getDateCellValue();
                    else
                        obj = cell.getNumericCellValue();;
                    break;
                case BOOLEAN:
                    obj = cell.getBooleanCellValue();
                    break;
                case BLANK:
                    break;
                case ERROR:
                    obj = cell.getErrorCellValue();
                    break;

                // CELL_TYPE_FORMULA will never occur
                case FORMULA:
                    obj = cell.getCellFormula();
                    break;
            }
        }
        return obj;
    }

    private Path saveToFile(MultipartFile file)
            throws IOException {
        if (file.isEmpty()) {
            throw new RuntimeException("Failed to store empty file");
        }

        try {
            String fileName = file.getOriginalFilename();
            InputStream is = file.getInputStream();

            Files.copy(is, Paths.get("upload/" + fileName),
                    StandardCopyOption.REPLACE_EXISTING);
            return Paths.get("upload/" + fileName);
        } catch (IOException e) {

            String msg = String.format("Failed to store file", file.getName());

            throw new RuntimeException(msg, e);
        }
    }
    /**
     * Creates a folder to desired location if it not already exists
     *
     * @param dirName
     *            - full path to the folder
     * @throws SecurityException
     *             - in case you don't have permission to create the folder
     */
    private void createFolderIfNotExists(String dirName)
            throws SecurityException {
        File theDir = new File(dirName);
        if (!theDir.exists()) {
            theDir.mkdir();
        }
    }


}
