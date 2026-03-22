@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo 企业级Excel导出工具 - 开始运行
echo ========================================
echo.

echo [1/3] 编译项目...
call mvn clean compile -DskipTests -q
if errorlevel 1 (
    echo 编译失败！
    pause
    exit /b 1
)

echo [2/3] 运行测试...
echo.
call mvn exec:java -Dexec.mainClass="com.example.excel.test.ExcelExportTest"
if errorlevel 1 (
    echo 运行失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 测试完成！请检查 D:\excelData 目录
echo ========================================
pause
