
options("repos" = c(CRAN="https://mirror.lzu.edu.cn/CRAN/"))
if(!require(openxlsx)){install.packages("openxlsx")}

###################################
###################################
####### create record file ########
###################################
###################################
# groupNum: 组别数
groupNum = 5
# repNum: 每个组别中重复样本数
repNum = 10
# weekNum: 预计记录周期
weekNum = 5
# FreqWeek：预计每周记录频次（默认每周记录两次）
FreqWeek = 2


# treatNum：预计所有处理类型的数目
# 当某个组别同时又不同处理条件时，使用该参数
# 该参数默认与 groupNum 数值相等。
treatNum = groupNum
#treatNum = 7

samples = groupNum * repNum
days = weekNum * FreqWeek

library(openxlsx)
wb <- openxlsx::createWorkbook()
sheet1 <- openxlsx::addWorksheet(wb, sheetName="BasicInfos", gridLines = FALSE)
sheet2 <- openxlsx::addWorksheet(wb, sheetName="DosingRecords", gridLines = FALSE)
sheet3 <- openxlsx::addWorksheet(wb, sheetName="SampleList", gridLines = FALSE)
sheet4 <- openxlsx::addWorksheet(wb, sheetName="DataRecords", gridLines = FALSE)

###########################
### Start of BasicInfos ###
###########################

## subject infos ##
subStyle <- openxlsx::createStyle(fontSize = 14, textDecoration = "bold", halign="left", valign="center")

openxlsx::mergeCells(wb, sheet1, cols=1:3, rows=2)
openxlsx::writeData(wb, sheet1, x="General infos", startCol=1, startRow = 2)
openxlsx::addStyle(wb, sheet1, subStyle, rows = 2, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet1, rows=2, heights=20)

for(i in 3:5){
	openxlsx::mergeCells(wb, sheet1, cols=1:2, rows=i)
	openxlsx::mergeCells(wb, sheet1, cols=3:21, rows=i)
}
subData1 <- data.frame(matrix(rep("", 63),nrow=3,ncol=21))
subData1$X1 = c("Study Number:", "Title:", "Objective:")
#subData1$X3 = c("1", "2", "3")
openxlsx::writeData(wb, sheet1, subData1, startCol = 1, startRow = 3, rowNames = FALSE, colNames = FALSE, borders = "surrounding", borderColour = "black")
textLeftAlign = openxlsx::createStyle(halign="right", valign="center")
openxlsx::addStyle(wb, sheet1, textLeftAlign, rows = 3:5, cols = 1, gridExpand = FALSE, stack = TRUE)

## personel ##
subStyle1 <- openxlsx::createStyle(fontSize = 12, textDecoration = "bold", halign="left", valign="center")
hs1 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","white","white"))
hs1add = openxlsx::createStyle(border=c("left","right"), borderColour="black")

openxlsx::mergeCells(wb, sheet1, cols=1:3, rows=7)
openxlsx::writeData(wb, sheet1, x="Personel", startCol=1, startRow = 7)
openxlsx::conditionalFormatting(wb, sheet1, cols=1, rows=7, rule = "!=0", style = subStyle1)
openxlsx::setRowHeights(wb, sheet1, rows=7, heights = 15)

for(ii in 8:13){
	openxlsx::mergeCells(wb, sheet1, cols=1:3, rows=ii)
	openxlsx::mergeCells(wb, sheet1, cols=4:8, rows=ii)
	openxlsx::mergeCells(wb, sheet1, cols=9:13, rows=ii)
}
subData2 <- data.frame(matrix(rep("", 65),nrow=5,ncol=13))
colnames(subData2)[1] = "Role"
colnames(subData2)[4] = "Name"
colnames(subData2)[9] = "Contact"
subData2$Role = c("Study Director", "Technician Group Leader", "Lab Technician", "Lab Technician", "Lab Technician")

openxlsx::writeData(wb, sheet1, subData2, startCol = 1, startRow = 8, rowNames = FALSE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::addStyle(wb, sheet1, createStyle(border=c("left"), borderColour="black"), rows = 8, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet1, createStyle(border=c("right"), borderColour="black"), rows = 8, cols = 13, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet1, createStyle(halign = "CENTER",valign = "CENTER"), rows = 9:13, cols = 2:13, gridExpand = TRUE, stack = TRUE)

## Treatments ##
openxlsx::mergeCells(wb, sheet1, cols=1:3, rows=15)
openxlsx::writeData(wb, sheet1, x="Treatments", startCol=1, startRow = 15)
openxlsx::conditionalFormatting(wb, sheet1, cols=1, rows=15, rule = "!=0", style = subStyle1)
openxlsx::setRowHeights(wb, sheet1, rows=7, heights = 15)

for(iii in 16:(16+treatNum)){
	openxlsx::mergeCells(wb, sheet1, cols=1:3, rows=iii)
	openxlsx::mergeCells(wb, sheet1, cols=4:13, rows=iii)
}
subData3 <- data.frame(matrix(rep("", treatNum*13),nrow=treatNum,ncol=13))
colnames(subData3)[1] = "Group"
colnames(subData3)[4] = "Treatment"
openxlsx::writeData(wb, sheet1, subData3, startCol = 1, startRow = 16, rowNames = FALSE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::addStyle(wb, sheet1, createStyle(border=c("left"), borderColour="black"), rows = 16, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet1, createStyle(border=c("right"), borderColour="black"), rows = 16, cols = 13, gridExpand = TRUE, stack = TRUE)

#########################
### End of BasicInfos ###
#########################

##################################
###  Start of Dosing Records   ###
##################################
subStyle3 <- openxlsx::createStyle(fontSize = 16, textDecoration = "bold", halign="center", valign="center")
subStyle4 <- openxlsx::createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3", halign="center", valign="center")

openxlsx::mergeCells(wb, sheet2, cols=1:8, rows=1)
openxlsx::writeData(wb, sheet2, x="Dosing record", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet2, subStyle3, rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet2, rows=1, heights=20)

openxlsx::mergeCells(wb, sheet2, cols=1:2, rows=2)
openxlsx::writeData(wb, sheet2, x="Study Number:", startCol=1, startRow = 2)
openxlsx::mergeCells(wb, sheet2, cols=3:10, rows=2)
openxlsx::writeFormula(wb, sheet2, x="=BasicInfos!C3", startCol=3, startRow = 2)

### legend ###
noteData = matrix(c("Note", "V", "dosed", "", "Note", "NA", "No dosing per protocol", "", "Note", "Holiday", "Holiday", ""), nrow=3, byrow=T)
openxlsx::mergeCells(wb, sheet2, cols=3:4, rows=3)
openxlsx::mergeCells(wb, sheet2, cols=3:4, rows=4)
openxlsx::mergeCells(wb, sheet2, cols=3:4, rows=5)
openxlsx::writeData(wb, sheet2, noteData, startCol=1, startRow = 3, colNames = FALSE, rowNames = FALSE, 
			borders = "columns", borderColour = "black", headerStyle = hs1)
openxlsx::mergeCells(wb, sheet2, cols=1, rows=3:5)
openxlsx::addStyle(wb, sheet2, createStyle(border="bottom"), rows = 3:4, cols = 1:4, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet2, createStyle(halign="center",valign="center",border="bottom"), rows = 3:5, cols = 1:2, gridExpand = TRUE, stack = TRUE)

openxlsx::setColWidths(wb, sheet2, cols = 6:(days+5), widths = 14)
openxlsx::mergeCells(wb, sheet2, cols=6:(days+5), rows=5)
openxlsx::writeData(wb, sheet2, x="Dates/Study Days", startCol=6, startRow = 5)
openxlsx::addStyle(wb, sheet2, subStyle4, rows = 5, cols = 6:(days+5), gridExpand = TRUE, stack = TRUE)
for(j in 1:days){
	#prjnt(j)
	openxlsx::writeData(wb, sheet2, x="Date:", startCol=(j-1)+6, startRow = 6)
	openxlsx::addStyle(wb, sheet2, createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3"), rows = 5, cols = (j-1)+6, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet2, x="Days:", startCol=(j-1)+6, startRow = 7)
	openxlsx::addStyle(wb, sheet2, createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3"), rows = 6, cols = (j-1)+6, gridExpand = TRUE, stack = TRUE)
}
subData4 <- data.frame(matrix(rep("", treatNum*repNum*days),nrow=treatNum*repNum,ncol=days))
openxlsx::writeData(wb, sheet2, subData4, startCol = 6, startRow = 8, rowNames = FALSE, colNames = FALSE, 
			borders = "rows", borderColour = "black", borderStyle="dashed")
openxlsx::addStyle(wb, sheet2, createStyle(halign="center", valign="center"), rows = 8:(8+treatNum*repNum - 1), cols = 6:(6+days - 1), gridExpand = TRUE, stack = TRUE)			
for(ji in 1:days){
	openxlsx::addStyle(wb, sheet2, createStyle(border="right"), rows = 8:(8+treatNum*repNum - 1), cols = 5+ji, gridExpand = TRUE, stack = TRUE)
}
for(jk in 1:treatNum){
	openxlsx::addStyle(wb, sheet2, createStyle(border="bottom"), rows = jk*repNum + 7, cols = 6:(6+days - 1), gridExpand = TRUE, stack = TRUE)
}

openxlsx::writeData(wb, sheet2, x="Group", startCol=1, startRow = 7)
openxlsx::addStyle(wb, sheet2, subStyle4, rows = 7, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::mergeCells(wb, sheet2, cols=2:4, rows=7)
openxlsx::writeData(wb, sheet2, x="Article", startCol=2, startRow = 7)
openxlsx::addStyle(wb, sheet2, subStyle4, rows = 7, cols = 2:4, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet2, x="SampleID", startCol=5, startRow = 7)
openxlsx::addStyle(wb, sheet2, subStyle4, rows = 7, cols = 5, gridExpand = TRUE, stack = TRUE)

openxlsx::setColWidths(wb, sheet2, cols = 1:5, widths = 13)
startRowID = 8
for(k in 1:treatNum){
	openxlsx::mergeCells(wb, sheet2, cols=1, rows=(startRowID+repNum*(k-1)):(startRowID+repNum*k-1))
	openxlsx::addStyle(wb, sheet2, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = (startRowID+repNum*(k-1)):(startRowID+repNum*k-1), cols = 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeFormula(wb, sheet2, x=paste0("=BasicInfos!A", k+16), startCol=1, startRow = startRowID+repNum*(k-1))
	openxlsx::mergeCells(wb, sheet2, cols=2:4, rows=(startRowID+repNum*(k-1)):(startRowID+repNum*k-1))
	openxlsx::addStyle(wb, sheet2, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center", wrapText=TRUE), rows = (startRowID+repNum*(k-1)):(startRowID+repNum*k-1), cols = 2:4, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeFormula(wb, sheet2, x=paste0("=BasicInfos!D", k+16), startCol=2, startRow = startRowID+repNum*(k-1))
	openxlsx::addStyle(wb, sheet2, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = (startRowID+repNum*(k-1)):(startRowID+repNum*k-1), cols = 5, gridExpand = TRUE, stack = TRUE)
	for(kk in 1:repNum){
		writeData(wb, sheet2, x="", startCol=5, startRow = startRowID+(k-1)*repNum+(kk-1))
	}
}

################################
###  End of Dosing Records   ###
################################

##################################
#####  Start of SampleList   #####
##################################

openxlsx::mergeCells(wb, sheet3, cols=1:8, rows=1)
openxlsx::writeData(wb, sheet3, x="Sample List", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet3, subStyle3, rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet3, rows=1, heights=20)

openxlsx::mergeCells(wb, sheet3, cols=1:2, rows=2)
openxlsx::writeData(wb, sheet3, x="Study Number:", startCol=1, startRow = 2)
openxlsx::mergeCells(wb, sheet3, cols=3:10, rows=2)
openxlsx::writeFormula(wb, sheet3, x="=BasicInfos!C3", startCol=3, startRow = 2)

### legend ###
noteData2 = matrix(c("Dissection", "Y", "Yes", "Dissection", "N", "No"), nrow=2, byrow=T)
openxlsx::writeData(wb, sheet3, noteData2, startCol=1, startRow = 3, colNames = FALSE, rowNames = FALSE, 
			borders = "columns", borderColour = "black", headerStyle = hs1)
openxlsx::mergeCells(wb, sheet3, cols=1, rows=3:4)
openxlsx::addStyle(wb, sheet3, createStyle(border="bottom"), rows = 3, cols = 1:3, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet3, createStyle(halign="center",valign="center",border="bottom"), rows = 3:4, cols = 1:3, gridExpand = TRUE, stack = TRUE)

headerList = c("Group", "SampleID", "Heart", "Liver", "Lung", "Kidney", "Stomach")
openxlsx::setColWidths(wb, sheet3, cols = 1:(length(headerList)+5), widths = 14)
for(hL in 1:length(headerList)){
	openxlsx::writeData(wb, sheet3, x=headerList[hL], startCol=hL, startRow = 7)
}
openxlsx::addStyle(wb, sheet3, subStyle4, rows = 7, cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)

subData42 <- data.frame(matrix(rep("", groupNum*repNum*(length(headerList)-2)),nrow=groupNum*repNum,ncol=length(headerList)-2))
openxlsx::writeData(wb, sheet3, subData42, startCol = 3, startRow = 8, rowNames = FALSE, colNames = FALSE, 
			borders = "rows", borderColour = "black", borderStyle="dashed")
openxlsx::addStyle(wb, sheet3, createStyle(halign="center", valign="center"), rows = 8:(8+groupNum*repNum - 1), cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)			
for(hi in 1:length(headerList)){
	openxlsx::addStyle(wb, sheet3, createStyle(border="right"), rows = 8:(8+groupNum*repNum - 1), cols = hi, gridExpand = TRUE, stack = TRUE)
}
for(hk in 1:groupNum){
	openxlsx::mergeCells(wb, sheet3, cols=1, rows=(8+repNum*(hk-1)):(8+repNum*hk-1))
	openxlsx::addStyle(wb, sheet3, createStyle(border="bottom"), rows = hk*repNum + 7, cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)
	openxlsx::writeFormula(wb, sheet3, x=paste0("=BasicInfos!A", hk+16), startCol=1, startRow = 8+repNum*(hk-1))
	for(hkk in 1:repNum){
		openxlsx::writeFormula(wb, sheet3, x=paste0("=DosingRecords!E", 8+(hk-1)*repNum+(hkk-1)), startCol=2, startRow = 8+(hk-1)*repNum+(hkk-1))
		openxlsx::addStyle(wb, sheet3, createStyle(border="bottom"), rows = 8+(hk-1)*repNum+(hkk-1), cols = 2, gridExpand = TRUE, stack = TRUE)
	}
}


################################
#####  End of SampleList   #####
################################

##################################
###  Start of Data Records   #####
##################################

openxlsx::mergeCells(wb, sheet4, cols=1:8, rows=1)
openxlsx::writeData(wb, sheet4, x="Data record", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet4, subStyle3, rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet4, rows=1, heights=20)

openxlsx::mergeCells(wb, sheet4, cols=1:2, rows=2)
openxlsx::writeData(wb, sheet4, x="Study Number:", startCol=1, startRow = 2)
openxlsx::mergeCells(wb, sheet4, cols=3:10, rows=2)
openxlsx::writeFormula(wb, sheet4, x="=BasicInfos!C3", startCol=3, startRow = 2)

openxlsx::setColWidths(wb, sheet4, cols = 1:(4*days+2), widths = 14)
for(m in 1:days){
	openxlsx::mergeCells(wb, sheet4, cols=((m-1)*4+3):(m*4+2), rows=4)
	openxlsx::mergeCells(wb, sheet4, cols=((m-1)*4+3):(m*4+2), rows=5)
	openxlsx::writeData(wb, sheet4, x="Dosing Date:", startCol=(m-1)*4+3, startRow = 4)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 4, cols = ((m-1)*4+3):(m*4+2), gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x="Days:", startCol=(m-1)*4+3, startRow = 5)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 5, cols = ((m-1)*4+3):(m*4+2), gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x="BodyWeight(g)", startCol=(m-1)*4+3, startRow = 6)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = (m-1)*4+3, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x="Length(mm)", startCol=(m-1)*4+4, startRow = 6)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = (m-1)*4+4, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x="Width(mm)", startCol=(m-1)*4+5, startRow = 6)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = (m-1)*4+5, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x="TV(mm3)", startCol=(m-1)*4+6, startRow = 6)
	openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = (m-1)*4+6, gridExpand = TRUE, stack = TRUE)
}
subData5 <- data.frame(matrix(rep("", groupNum*repNum*days*4),nrow=groupNum*repNum,ncol=days*4))
openxlsx::writeData(wb, sheet4, subData5, startCol = 3, startRow = 7, rowNames = FALSE, colNames = FALSE, 
			borders = "all", borderColour = "black")
openxlsx::addStyle(wb, sheet4, createStyle(halign="center", valign="center"), rows = 7:(7+groupNum*repNum - 1), cols = 3:(3+days*4 - 1), gridExpand = TRUE, stack = TRUE)	
colList <- toupper(letters)
for(s in 1:length(letters)){
	for(t in 1:length(letters)){
		colList <- c(colList, paste0(toupper(letters)[s],toupper(letters)[t]))
	}
}
for(p in 1:days){
	for(q in 7:(7+groupNum*repNum - 1)){
		preOne = (p-1)*4+6 - 2
		preTwo = (p-1)*4+6 - 1
		openxlsx::writeFormula(wb, sheet4, x=paste0("=",colList[preOne],q,"*",colList[preTwo],q,"*",colList[preTwo],q,"/2"), startCol=(p-1)*4+6, startRow = q)
	}
}

openxlsx::writeData(wb, sheet4, x="Group", startCol=1, startRow = 6)
openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet4, x="SampleID", startCol=2, startRow = 6)
openxlsx::addStyle(wb, sheet4, subStyle4, rows = 6, cols = 2, gridExpand = TRUE, stack = TRUE)
startRowID = 7
for(n in 1:groupNum){
	openxlsx::mergeCells(wb, sheet4, cols=1, rows=(startRowID+repNum*(n-1)):(startRowID+repNum*n-1))
	openxlsx::addStyle(wb, sheet4, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = (startRowID+repNum*(n-1)):(startRowID+repNum*n-1), cols = 1, gridExpand = TRUE, stack = TRUE)
	writeFormula(wb, sheet4, x=paste0("=BasicInfos!A", n+16), startCol=1, startRow = startRowID+repNum*(n-1))
	for(nn in 1:repNum){
		writeFormula(wb, sheet4, x=paste0("=DosingRecords!E", startRowID+1+(n-1)*repNum+(nn-1)), startCol=2, startRow = startRowID+(n-1)*repNum+(nn-1))
	}
	openxlsx::addStyle(wb, sheet4, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = (startRowID+repNum*(n-1)):(startRowID+repNum*n-1), cols = 2, gridExpand = TRUE, stack = TRUE)
}

################################
###  End of Data Records   #####
################################

openxlsx::saveWorkbook(wb, "newRecord.xlsx", overwrite = TRUE)

###################################
###################################
####### create record file ########
###################################
###################################

