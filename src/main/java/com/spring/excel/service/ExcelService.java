package com.spring.excel.service;

import com.spring.excel.helper.ExcelHelper;
import com.spring.excel.model.Tutorial;
import com.spring.excel.repository.TutorialRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ExcelService {
    @Autowired
    private TutorialRepository tutorialRepository;

    public void save(MultipartFile file){
        try{
            List<Tutorial> tutorials= ExcelHelper.excelToTutorials(file.getInputStream());
            tutorialRepository.saveAll(tutorials);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: "+ e.getMessage());
        }
    }

    public List<Tutorial> getAllTutorials(){
        return tutorialRepository.findAll();
    }
    
    
    public ByteArrayInputStream dataToExcel() throws Exception {
    return ExcelHelper.tutorialToExcel(tutorialRepository.findAll());
    }

}
