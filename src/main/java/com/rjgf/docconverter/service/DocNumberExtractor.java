package com.rjgf.docconverter.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.nio.charset.StandardCharsets;

public class DocNumberExtractor {
    private final HWPFDocument document;
    private final int[] counters;
    private int lastLevel;
    private static final int MAX_PARAGRAPH_GAP = 6; // 定义最大允许的段落间隔
    private int lastParagraphNumber = -1; // 记录上一个编号段落的位置
    private int headerColumnCount = -1; // 添加成员变量存储表头列数
    // private int currentTableStartOffset = -1;  // 记录当前表格的起始位置

    private boolean isFirstCellInTable = false;

    public DocNumberExtractor(HWPFDocument document) {
        this.document = document;
        this.counters = new int[10];
        this.lastLevel = 0;
    }

    public static String extractFromFile(InputStream inputStream) {
        try (HWPFDocument document = new HWPFDocument(inputStream)) {
            DocNumberExtractor extractor = new DocNumberExtractor(document);
            return extractor.extract();
        } catch (IOException e) {
            e.printStackTrace();
            return "";
        }
    }

    public String extract() {
        StringBuilder formattedText = new StringBuilder();
        Range range = document.getRange();

        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph para = range.getParagraph(i);
            String text = para.text().trim();

            // if (!text.isEmpty()) {
                if (para.isInTable()) {
                    processTableCell(formattedText, para);
                } else if (para.isInList()) {
                    // 检查段落间隔
                    if (shouldResetCounters(i)) {
                        resetCounters();
                    }
                    processListItem(formattedText, para, text);
                    lastParagraphNumber = i;
                } else {
                    processNormalParagraph(formattedText, para);
                }
            // }
        }

        String content = formattedText.toString();

        content = content.replace("&&||", "&&");

        return content;
    }

    private void processTableCell(StringBuilder formattedText, Paragraph para) {
        try {
            String cellText = para.text();

            // 处理首个单元格 - 只在新表格开始时添加Table:标记
            if (!isFirstCellInTable) {
                // 检查是否是新表格
                if (isFirstCellInTable(para)) {
                    formattedText.append("\nTable:");
                    isFirstCellInTable = true;
                }
            } 

            boolean hasNewline = cellText.contains("\n") || cellText.contains("\r");

            // 处理单元格内容
            cellText = processCellContent(cellText);
            
            // 如果包含换行，且在单元格首次出现，则添加标题，否则添加分隔符
            if (hasNewline) {
                formattedText.append("||").append(cellText).append("&&");
            } else {
                // 如果是最后一个值则去掉最后一个"|"
                formattedText.append("||").append(cellText);
            }
            // 处理行尾
            if (isLastCellInRow(para)) {
                // 如果是第一行（表头），记录列数
                if (headerColumnCount == -1) {
                    headerColumnCount = countCellsInCurrentRow(formattedText.toString());
                }     
                // 使用表头的列数来填充剩余单元格
                int currentCells = countCellsInCurrentRow(formattedText.toString());
                for (int i = currentCells; i < headerColumnCount; i++) {
                    formattedText.append("||").append("-");
                }
                // 只有在不是表格最后一行时才添加 +++
                if (!isLastRowInTable(para)) {
                    formattedText.append("+++");
                } else {
                    isFirstCellInTable = false;
                }
                formattedText.append("");
            }
        } catch (Exception e) {
            formattedText.append("||ERROR||+++\n");
        }
    }

    private String processCellContent(String cellText) {
        if (cellText == null) {
            return "";
        }
        
        // 去除首尾空白
        cellText = cellText.trim();
        if (cellText.isEmpty()) {
            return "";
        }

        // System.out.println("cellText: " + cellText);
        
        // 处理换行: 将换行符替换为 <br> 标记
        cellText = cellText.replaceAll("\\r?\\n", "<br>");
        
        // 处理连续空格
        cellText = cellText.replaceAll("\\s+", " ");
        
        // 处理可能导致格式混乱的字符
        cellText = cellText.replace("||", "│"); // 替换可能干扰表格格式的分隔符
        cellText = cellText.replace("+++", "＋"); // 替换可能干扰行结束符的字符
        
        return cellText;
    }

    private boolean isFirstCellInTable(Paragraph para) {
        try {
            // 确保段落在表格中
            if (!para.isInTable()) {
                return false;
            }
            
            // 获取文档范围
            Range range = document.getRange();
            int currentIndex = -1;
            
            // 找到当前段落的索引
            for (int i = 0; i < range.numParagraphs(); i++) {
                if (range.getParagraph(i).getStartOffset() == para.getStartOffset()) {
                    currentIndex = i;
                    break;
                }
            }
            
            // 如果找不到当前段落，返回false
            if (currentIndex == -1) {
                return false;
            }
            
            // 如果是第一个段落，且在表格中，则是表格第一个单元格
            if (currentIndex == 0) {
                return true;
            }
            
            // 检查前一个段落是否不在表格中
            Paragraph prevPara = range.getParagraph(currentIndex - 1);
            return !prevPara.isInTable();
            
        } catch (Exception e) {
            System.err.println("Error in isFirstCellInTable: " + e.getMessage());
            return false;
        }
    }

    private boolean isLastCellInRow(Paragraph para) {
        try {
            return para.isTableRowEnd();
        } catch (Exception e) {
            return false;
        }
    }

    private int countCellsInCurrentRow(String text) {
        int lastNewlineIndex = text.lastIndexOf('\n');
        String currentRow = text.substring(lastNewlineIndex + 1);
        return (int) currentRow.chars().filter(ch -> ch == '|').count() / 2;
    }

    private void processListItem(StringBuilder formattedText, Paragraph para, String text) {
        int level = para.getIlvl() + 1;
        updateCounters(level);
        String number = generateNumber(level);
        formattedText.append(formatMarkdown(number, text, level));
        lastLevel = level;
    }

    private void updateCounters(int currentLevel) {
        if (currentLevel <= lastLevel) {
            for (int j = currentLevel; j < counters.length; j++) {
                counters[j] = 0;
            }
        }
        counters[currentLevel - 1]++;
    }

    private String generateNumber(int level) {
        StringBuilder number = new StringBuilder();
        for (int j = 0; j < level; j++) {
            if (j > 0) number.append(".");
            number.append(counters[j]);
        }
        return number.toString();
    }

    private void processNormalParagraph(StringBuilder formattedText, Paragraph para) {
        String text = para.text().trim();
        
        // 如果需要添加分隔符
        if (shouldAddSeparator(para, text)) {
            formattedText.append("+".repeat(50)).append("\n");
        }
        
        formattedText.append(text).append("\n\n");
    }

    private boolean shouldAddSeparator(Paragraph para, String text) {
        if (text.isEmpty()) {
            return false;
        }

        // 检查文本长度，标题通常不会太长
        if (text.length() > 50) {  // 增加长度限制，允许更长的标题
            return false;
        }

        // 检查是否为标准标题格式（第X章/部分）
        String titlePattern = "^\\s*第[一二三四五六七八九十\\d]{1,3}(章|部分|节|条|款|项).*$";
        boolean isStandardTitle = text.matches(titlePattern);

        // 检查结束标记
        String[] endingMarks = {"。", "！", "!", "？", "?", ";", "；", ":", "：", "...", "★"};
        String trimmedText = text.trim();
        boolean hasEndingMark = Arrays.stream(endingMarks)
                .anyMatch(trimmedText::endsWith);

        // 检查样式和对齐方式
        int leftIndent = para.getIndentFromLeft();
        int firstLineIndent = para.getFirstLineIndent();
        boolean isCentered = para.getJustification() == 1; // 1表示居中
        
        // 判断是否符合标题格式
        boolean hasSpecialFormatting = isCentered || leftIndent > 1000000 || firstLineIndent > 1000000;

        // 如果是标准标题格式，或者具有特殊格式（居中/缩进）且没有结束标记
        return (isStandardTitle || hasSpecialFormatting) && !hasEndingMark;
    }

    private String formatMarkdown(String number, String text, int level) {
        // String prefix = "#".repeat(level);
        return " Pnumber " + number + " " + text + ":\n\n";
    }

    private boolean shouldResetCounters(int currentParagraphNumber) {
        // 如果是第一个编号段落，不需要重置
        if (lastParagraphNumber == -1) {
            return false;
        }
        
        // 如果段落间隔超过设定值，重置计数器
        return (currentParagraphNumber - lastParagraphNumber) > MAX_PARAGRAPH_GAP;
    }

    private void resetCounters() {
        // 重置所有计数器
        Arrays.fill(counters, 0);
        lastLevel = 0;
    }

    private boolean isLastRowInTable(Paragraph para) {
        try {
            // 获取下一个段落
            Range range = document.getRange();
            int currentIndex = -1;
            
            // 找到当前段落的索引
            for (int i = 0; i < range.numParagraphs(); i++) {
                if (range.getParagraph(i).getStartOffset() == para.getStartOffset()) {
                    currentIndex = i;
                    break;
                }
            }
            
            if (currentIndex == -1 || currentIndex >= range.numParagraphs() - 1) {
                return true;
            }
            
            // 检查下一个段落是否还在表格中
            Paragraph nextPara = range.getParagraph(currentIndex + 1);
            return !nextPara.isInTable();
            
        } catch (Exception e) {
            System.err.println("Error in isLastRowInTable: " + e.getMessage());
            return false;
        }
    }

    // main方法
    public static void main(String[] args) {
        String filePath = "test.doc";
        try {
            File file = new File(filePath);
            String extractedText = DocNumberExtractor.extractFromFile(new FileInputStream(file));
            // 输出到文件 乱码
            // 使用UTF-8编码输出
            FileWriter writer = new FileWriter("output.txt", StandardCharsets.UTF_8);
            writer.write(extractedText);
            writer.close();
            // System.out.println(extractedText);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
