package com.rjgf.docconverter.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

public class DocNumberExtractor {
    private final HWPFDocument document;
    private final int[] counters;
    private int lastLevel;
    private static final int MAX_PARAGRAPH_GAP = 6; // 定义最大允许的段落间隔
    private int lastParagraphNumber = -1; // 记录上一个编号段落的位置

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

        return formattedText.toString();
    }

    private void processTableCell(StringBuilder formattedText, Paragraph para) {
        try {
            String cellText = para.text();
            
            // 处理首个单元格
            if (isFirstCellInTable(para)) {
                formattedText.append("\nTable:\n");
            }
            
            // 处理单元格内容
            cellText = processCellContent(cellText);
            
            // 始终添加单元格分隔符和内容
            formattedText.append("||").append(cellText);
            
            // 处理行尾
            if (isLastCellInRow(para)) {
                // 确保行中的所有单元格都被填充
                int expectedCells = 10; // 假设最大列数为10
                int currentCells = countCellsInCurrentRow(formattedText.toString());
                for (int i = currentCells; i < expectedCells; i++) {
                    formattedText.append("||").append("-");
                }
                formattedText.append("||+++\n"); // 使用 +++ 作为行结束符
            }
        } catch (Exception e) {
            formattedText.append("||ERROR||+++\n");
        }
    }

    private String processCellContent(String cellText) {
        if (cellText == null) {
            return "-";  // 空值用 - 代替
        }
        
        // 去除首尾空白
        cellText = cellText.trim();
        if (cellText.isEmpty()) {
            return "-";
        }
        
        // 处理换行: 将换行符替换为特殊标记
        cellText = cellText.replaceAll("\\r?\\n", "<br>");
        
        // 处理连续空格
        cellText = cellText.replaceAll("\\s+", " ");
        
        return cellText;
    }

    private boolean isFirstCellInTable(Paragraph para) {
        try {
            return para.getTableLevel() > 0 && 
                   (para.getStartOffset() == para.getTable(para).getStartOffset());
        } catch (Exception e) {
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
        if (text.length() > 20) {
            return false;
        }

        // 检查是否匹配"第X章/部分"模式
        String pattern = "^\\s*第[一二三四五六七八九十\\d]{1,3}(章|部分)[节]?.*$";
        if (!text.matches(pattern)) {
            return false;
        }

        // 检查结束标记
        String[] endingMarks = {"。", "！", "!", "？", "?", ";", "；", ":", "：", "...", "★"};
        String trimmedText = text.trim();
        for (String mark : endingMarks) {
            if (trimmedText.endsWith(mark)) {
                return false;
            }
        }

        // 检查最后一个字符是否为数字
        if (Character.isDigit(trimmedText.charAt(trimmedText.length() - 1))) {
            return false;
        }

        // 检查样式和对齐方式
        boolean isNormalOrText = para.getStyleIndex() == 0; // 假设0是Normal样式
        int leftIndent = para.getIndentFromLeft();
        boolean isCentered = para.getJustification() == 1; // 1通常表示居中

        return !isNormalOrText || (isNormalOrText && (leftIndent > 1000000 || isCentered));
    }

    private String formatMarkdown(String number, String text, int level) {
        // String prefix = "#".repeat(level);
        return " " + number + " " + text + "\n\n";
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
}
