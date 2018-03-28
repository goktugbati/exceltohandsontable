package com.excel.demo.controller;


import com.excel.demo.model.SheetModel;
import com.excel.demo.service.ExcelPOIHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
public class MainController {

    private String fileLocation;
    private final ExcelPOIHelper excelPOIHelper;

    @Autowired
    public MainController(ExcelPOIHelper excelPOIHelper) {
        this.excelPOIHelper = excelPOIHelper;
    }


    @GetMapping("/")
    public String home(){
        return "excel_upload";
    }

    @PostMapping("/uploadExcelFile")
    public String uploadFile(Model model, MultipartFile file) throws IOException {
        InputStream in = file.getInputStream();
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        fileLocation = path.substring(0, path.length() - 1) + file.getOriginalFilename();
        FileOutputStream f = new FileOutputStream(fileLocation);
        int ch = 0;
        while ((ch = in.read()) != -1) {
            f.write(ch);
        }
        f.flush();
        f.close();
        model.addAttribute("message", "File: " + file.getOriginalFilename()
                + " has been uploaded successfully!");
        return "redirect:/sheetnames";
    }

    //User selects the sheet they'd like to work on
    @GetMapping(value = "/sheetnames", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public String sheetNamesList(Model model) throws IOException {
        List<SheetModel> sheetModelList = excelPOIHelper.getSheetNames(fileLocation);
        model.addAttribute("sheetmodels", sheetModelList);

        return "excel_sheet";
    }

    //Displays the chosen sheet on a table that allows user to make changes on the excel file
    @PostMapping(value = "/readPOI", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public String readPOI(@RequestParam int sheetNo, Model model) throws IOException {

        if (fileLocation != null) {
            if (fileLocation.endsWith(".xlsx") || fileLocation.endsWith(".xls")) {

                Map<Integer, List<String>> data
                        = excelPOIHelper.readMyExcelFileLikeaBOSS(fileLocation, sheetNo);

                List<String> columnNames = data.get(0);
                List<List<String>> dataList = new ArrayList<>();
                data.remove(0);
                for(int i = 1; i<data.size(); i++){
                    dataList.add(data.get(i));
                }
                model.addAttribute("data", dataList);
                model.addAttribute("columnNames", columnNames);
                model.addAttribute("sheetNo", sheetNo);
            } else {
                model.addAttribute("message", "Not a valid excel file!");
            }
        } else {
            model.addAttribute("message", "File missing! Please upload an excel file.");
        }
        return "excel";
    }

    @PostMapping(value = "/updateexcel")
    public void updateExcel(@RequestParam int sheetNo,
                              @RequestParam int rowIndex,
                              @RequestParam int cellIndex,
                              @RequestParam String cellContent){
        try {
            excelPOIHelper.updateExcelFile(fileLocation, sheetNo, cellContent, rowIndex, cellIndex);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }



}
