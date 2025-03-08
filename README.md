# Doc Converter

一个基于Spring Boot的文档转换服务,可以将Word文档(.doc)转换为纯文本格式。

## 功能特性

- 支持.doc格式的Word文档转换
- 转换后保留文档的段落结构
- RESTful API接口
- 文件上传和下载支持

## 快速开始

### 环境要求

- JDK 8+
- Maven 3.6+
- Spring Boot 2.x

### 运行应用

1. 克隆项目 `git clone https://github.com/xla145/doc-converter.git`
2. 运行 `mvn spring-boot:run` 启动应用
3. 访问 `http://localhost:8080/api/convert/doc-to-text` 进行文档转换

### 文档转换

1. 上传Word文档
2. 点击转换按钮
3. 下载转换后的纯文本文件

### 注意事项

- 仅支持.doc文件
- 转换后的文本会保留段落结构，但可能会包含一些格式化问题
- 转换后的文本会以附件形式下载