@echo off
chcp 65001
cd /d C:\Users\lance\IdeaProjects\excelUitl
call mvn clean compile exec:java -Dexec.mainClass=com.example.excel.test.ExcelExportTest
