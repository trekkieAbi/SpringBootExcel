package com.spring.excel.controller;

import com.spring.excel.helper.ExcelHelper;
import com.spring.excel.message.ResponseMessage;
import com.spring.excel.model.Tutorial;
import com.spring.excel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@CrossOrigin("http://localhost:8081")
@Controller
@RequestMapping("/api/excel")
public class ExcelController {
    @Autowired
    ExcelService excelService;

    @PostMapping("/upload")
    public ResponseEntity<ResponseMessage> uploadFile(@RequestParam("file") MultipartFile file) {
        String message = "";
        if (ExcelHelper.hasExcelFormat(file)) {
            try {
                excelService.save(file);
                message = "Uplaoded the file successfully " + file.getOriginalFilename();
                return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
            } catch (Exception e) {
                message = "Could not upload the file: " + file.getOriginalFilename();
                return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
            }
        }

    message="Please upload an excel file!";
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
}

@GetMapping("/tutorials")
public ResponseEntity<List<Tutorial>> getAllTutorials(){
        try{
            List<Tutorial> tutorials=excelService.getAllTutorials();
            if(tutorials.isEmpty()){
                return new ResponseEntity<>(HttpStatus.NO_CONTENT);
            }
            return new ResponseEntity<>(tutorials,HttpStatus.OK);
        }catch (Exception e){
            return new ResponseEntity<>(null,HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }


}



