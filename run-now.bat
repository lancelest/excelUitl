@echo off
chcp 65001
cd /d C:\Users\lance\IdeaProjects\excelUitl
echo Compiling...
call mvn clean compile -DskipTests -q
if errorlevel 1 (
    echo Compile failed!
    exit /b 1
)
echo Running...
call mvn exec:java -Dexec.mainClass=com.example.excel.test.ExcelExportTest -q
echo Done.
