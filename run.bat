@echo off
chcp 65001

cd /d C:\Users\lance\IdeaProjects\excelUitl

set CP=target\classes
set CP=%CP%;E:\MavenRepository\org\apache\poi\poi\4.1.2\poi-4.1.2.jar
set CP=%CP%;E:\MavenRepository\commons-codec\commons-codec\1.13\commons-codec-1.13.jar
set CP=%CP%;E:\MavenRepository\org\apache\commons\commons-math3\3.6.1\commons-math3-3.6.1.jar
set CP=%CP%;E:\MavenRepository\com\zaxxer\SparseBitSet\1.2\SparseBitSet-1.2.jar
set CP=%CP%;E:\MavenRepository\org\apache\poi\poi-ooxml\4.1.2\poi-ooxml-4.1.2.jar
set CP=%CP%;E:\MavenRepository\org\apache\poi\poi-ooxml-schemas\4.1.2\poi-ooxml-schemas-4.1.2.jar
set CP=%CP%;E:\MavenRepository\org\apache\xmlbeans\xmlbeans\3.1.0\xmlbeans-3.1.0.jar
set CP=%CP%;E:\MavenRepository\org\apache\commons\commons-compress\1.19\commons-compress-1.19.jar
set CP=%CP%;E:\MavenRepository\com\github\virtuald\curvesapi\1.06\curvesapi-1.06.jar
set CP=%CP%;E:\MavenRepository\org\projectlombok\lombok\1.18.30\lombok-1.18.30.jar
set CP=%CP%;E:\MavenRepository\org\apache\commons\commons-lang3\3.14.0\commons-lang3-3.14.0.jar
set CP=%CP%;E:\MavenRepository\org\apache\commons\commons-collections4\4.4\commons-collections4-4.4.jar

java -cp "%CP%" com.example.excel.test.ExcelExportTest
