@echo off
chcp 65001
cd /d C:\Users\lance\IdeaProjects\excelUitl
mvn exec:java -Dexec.mainClass=com.example.excel.test.ExcelExportTest -q 2>&1
