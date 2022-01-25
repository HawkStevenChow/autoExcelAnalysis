# autoExcelAnalysis
create new Excel file with openxlsx package in R, then collect informations from existed Excel file and do statistics analysis

1. Files
   1) "CreateNewXlsx.R": create new Excel file with multiple sheets, and format font, size, background of sheets;
   2) "record22.xlsx": an example of recorded file with experimental data; 
   3) "AnalysisXlsx4Results.R": analysis results of recorded file, with statistical tables and plots;
   4) "testCreateXlsx.bat": windows bat file for "CreateNewXlsx.R";
   5) "testAnalysisXlsx.bat": windows bat file for "AnalysisXlsx4Results.R";
   
2. Usages
	1)创建数据录入模板文件(如，“数据录入.xlsx”)；
		步骤一：使用“notepad++”打开文件“CreateNewXlsx.R”，设定“组别数”，如实验设计有5组，则 groupNum = 5;
				设定“每个组别中重复样本数”，如每组包含10个重复，则设置 repNum = 10;
				设定“预计记录周期”，如预计记录5周，则设置 weekNum = 5;
				设定“预计每周记录频次”，如预计每周记录两次，则设置 FreqWeek = 2;
				'Ctrl+S'，保存修改好的文件；
		步骤二：双击“testCreateXlsx.bat”，脚本会根据步骤一中的参数设置，生成相应的数据录入文件，
				如，“数据录入.xlsx”
	
	2)数据分析和整理
		步骤一：将记录有实验数据的文件“数据录入.xlsx”，重命名为“recorded22.xlsx”
				(或者修改文件“AnalysisXlsx4Results.R”中EXCEL文件名称),保存修改好的文件；
		步骤二：双击“testAnalysisXlsx.bat”，脚本会自动读入数据记录文件，进行数据分析，并将分析结果
				整理并保存到结果文件“analysis-result.xlsx”中
  
