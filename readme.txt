////////////////////////////////////////////////////////////////
复制并导入文件到Excel表格的VBA代码

文件说明：
create_folder_copy_and_import.xlsm 包含宏（VBA代码），可以复制文件到指定目录（在Excel表格中指定，如果目录不存在则创建），并导入文本文件到Excel表格。

create_folder_and_copy_file.xlsm 的内容：
第一列是待复制的文件，第二列是目标子目录，第三列是待复制的空白worksheet，第四列是导入.txt文件时给sheet命名的前缀。

AIB_20131030.txt, BOI_20131030.txt, AIB_20131107.TXT, BOI_20131107.TXT 是数据文件。
worksheet.xlsx 是空白worksheet，用于导入.txt文件。

使用方法：
下载整个项目，然后使用Excel 2010打开create_folder_and_copy_file.xlsm，启用并运行宏，则该文件所在目录下的文件就会根据Excel表格中指定的规则（表格第一列为源路径（可使用通配符如*.txt），第二列为目标路径）被复制到目标目录中。

////////////////////////////////////////////////////////////////
