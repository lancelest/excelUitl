@echo off
chcp 65001
cd /d C:\Users\lance\IdeaProjects\excelUitl
mvn compile exec:java -Dexec.mainClass=com.example.excel.test.ExcelExportTest -DskipTests 2>&1
