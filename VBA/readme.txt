文件说明：
Module1.bas	是VBA代码。
Worksheet1.xlsm	是包含宏的Excel文件，Worksheet1.xlsx不包含宏。
list.txt	是文件列表
AIB_20131030.txt, BOI_20131030.txt, Ulster_20131030.txt 是数据文件。

使用说明：
用Excel 2010打开Worksheet1.xlsm，启用宏。
选中Sheet1中的文件列表（或者选中其中文件列表其中一个格也可以）。
运行宏，就会为每个文件创建新的tab，读入文件内容，并将tab改名为文件名+当前时间。
