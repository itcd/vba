03 April 2014

Whats New:
1. Improve user experience, more user friendly interface
2. The macro will now move the data file instead of copying them into each folder. So the user could know if any data file was unused by the Macro.
3. The Completed Worksheet.xlsx file will now automatically change its name by on the date.
eg. in the 20130203 folder, the worksheet name will be xxxx as at 20130203.xlsx (the worksheet.xlsx is now called All segreggated bank files as at.xlsx)

////////////////////////////////////////////////////////////////

���Ʋ������ļ���Excel����VBA����

�ļ�˵����
create_folder_copy_and_import.xlsm �����꣨VBA���룩�����Ը����ļ���ָ��Ŀ¼����Excel�����ָ�������Ŀ¼�������򴴽������������ı��ļ���Excel���
worksheet.xlsx �ǿհ�worksheet�����ڵ���.txt�ļ���
AIB_20131030.txt, BOI_20131030.txt, AIB_20131107.TXT, BOI_20131107.TXT �������ļ���

�ļ���ʽ��ʾ����
create_folder_and_copy_file.xlsm ���ݵĸ�ʽ���£�
��һ���Ǵ����Ƶ��ļ���
�ڶ�����Ŀ����Ŀ¼��
�������Ǵ����ƵĿհ�worksheet��
��������reference number������.txt�ļ�ʱ������sheet������ǰ׺��

����create_folder_copy_and_import.xlsm ��һ����
AIB*.txt	AIB	worksheet.xlsx	ref0001

VBA����ͻὫ��ǰĿ¼�����з���AIB*.txt���ļ��Լ�worksheet.xlsx���Ƶ���ΪAIB����Ŀ¼�������Ŀ¼�������򴴽�����
Ȼ��AIB�����Ŀ¼�µ�����txt�ļ���AIB\*.txt�����뵽�ղŸ��Ƶ�AIBĿ¼�µ�worksheet.xlsx��AIB\worksheet.xlsx����
�����ʱ��ÿ��tab��������ref0001����.txt���ļ�����

ʹ�÷�����
1. ����������Ŀ��Ȼ��ʹ�� Excel 2010 �� create_folder_and_copy_file.xlsm;
2. ���ò����к꣬����ļ�����Ŀ¼�µ��ļ��ͻ����Excel�����ָ���Ĺ��򣨱���һ��ΪԴ·������ʹ��ͨ�����*.txt����
�ڶ���ΪĿ��·���������Ƶ�Ŀ��Ŀ¼�С�

////////////////////////////////////////////////////////////////
