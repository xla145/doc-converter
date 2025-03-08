package com.rjgf.docconverter.service;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@Service
public class DocConverterService {
    
    public String convertDocToText(MultipartFile file) throws IOException {
        String fileName = file.getOriginalFilename();
        InputStream inputStream = file.getInputStream();

        if (fileName == null) {
            throw new IllegalArgumentException("文件名不能为空");
        }
        
        if (fileName.endsWith(".doc")) {
            return processDoc(inputStream);
        } else {
            throw new IllegalArgumentException("不支持的文件格式，仅支持 .doc 文件");
        }
    }

    private String processDoc(InputStream inputStream) throws IOException {
        String extractor = DocNumberExtractor.extractFromFile(inputStream);
        return extractor;
    }

    public void saveToFile(String content, String outputPath) throws IOException {
        try (java.io.FileWriter writer = new java.io.FileWriter(outputPath)) {
            writer.write(content);
        }
    }
} 