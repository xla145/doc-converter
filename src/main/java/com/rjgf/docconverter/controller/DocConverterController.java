package com.rjgf.docconverter.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import com.rjgf.docconverter.service.DocConverterService;

import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/api/convert")
public class DocConverterController {

    @Autowired
    private DocConverterService docConverterService;

    @PostMapping(value = "/doc-to-text")
    public ResponseEntity<?> convertDocToText(@RequestParam("file") MultipartFile file) {
        try {
            String text = docConverterService.convertDocToText(file);
            
            // 创建文件名
            String fileName = file.getOriginalFilename() + ".txt";
            
            // 创建响应头
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.TEXT_PLAIN);
            headers.setContentDispositionFormData("attachment", fileName);
            
            // 将文本转换为字节流
            byte[] documentBody = text.getBytes(StandardCharsets.UTF_8);
            
            return ResponseEntity
                    .ok()
                    .headers(headers)
                    .contentLength(documentBody.length)
                    .body(documentBody);
                    
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("status", "error");
            error.put("message", "转换失败：" + e.getMessage());
            return ResponseEntity.badRequest().body(error);
        }
    }
} 