03 April 2014

Whats New:
1. Improve user experience, more user friendly interface
2. The macro will now move the data file instead of copying them into each folder. So the user could know if any data file was unused by the Macro.
3. The Completed Worksheet.xlsx file will now automatically change its name by on the date.
eg. in the 20130203 folder, the worksheet name will be xxxx as at 20130203.xlsx (the worksheet.xlsx is now called All segreggated bank files as at.xlsx)

////////////////////////////////////////////////////////////////

复制并导入文件到Excel表格的VBA代码

文件说明：
create_folder_copy_and_import.xlsm 包含宏（VBA代码），可以复制文件到指定目录（在Excel表格中指定，如果目录不存在则创建），并导入文本文件到Excel表格。
worksheet.xlsx 是空白worksheet，用于导入.txt文件。
AIB_20131030.txt, BOI_20131030.txt, AIB_20131107.TXT, BOI_20131107.TXT 是数据文件。

文件格式与示例：
create_folder_and_copy_file.xlsm 内容的格式如下：
第一列是待复制的文件；
第二列是目标子目录；
第三列是待复制的空白worksheet；
第四列是reference number，导入.txt文件时用作给sheet命名的前缀。

例如create_folder_copy_and_import.xlsm 第一行是
AIB*.txt	AIB	worksheet.xlsx	ref0001

VBA代码就会将当前目录下所有符合AIB*.txt的文件以及worksheet.xlsx复制到名为AIB的子目录（如果子目录不存在则创建），
然后将AIB这个子目录下的所有txt文件（AIB\*.txt）导入到刚才复制到AIB目录下的worksheet.xlsx（AIB\worksheet.xlsx），
导入的时候每个tab的名字是ref0001加上.txt的文件名。

使用方法：
1. 下载整个项目，然后使用 Excel 2010 打开 create_folder_and_copy_file.xlsm;
2. 启用并运行宏，则该文件所在目录下的文件就会根据Excel表格中指定的规则（表格第一列为源路径（可使用通配符如*.txt），
第二列为目标路径）被复制到目标目录中。

////////////////////////////////////////////////////////////////
