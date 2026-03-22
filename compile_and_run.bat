@echo off
chcp 65001
cd /d C:\Users\lance\IdeaProjects\excelUitl
echo Compiling...
call mvn compile -DskipTests
echo Running...
call mvn exec:java -Dexec.mainClass=com.example.excel.test.ExcelExportTest
