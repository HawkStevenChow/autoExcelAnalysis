options("repos" = c(CRAN="https://mirror.lzu.edu.cn/CRAN/"))
if(!require(xlsx)){install.packages("xlsx")}

# groupNum: 实验组别数
groupNum = 5
# repNum: 每个组别中重复样本数
repNum = 10
# weekNum: 预计记录周期
weekNum = 5
# DaysPerWeek : 每周记录天数
DaysPerWeek = 2

samples = groupNum * repNum
days = weekNum * DaysPerWeek

library(xlsx)
wb <- createWorkbook()
sheet1 <- createSheet(wb, sheetName="基本信息")
sheet2 <- createSheet(wb, sheetName="个体数据")


cb1 <- CellBlock(sheet1, 2, 1, samples + groupNum + 20, 8+days+3)
cs1 <- CellStyle(wb) + Alignment(h = "ALIGN_CENTER", v="VERTICAL_CENTER") + Fill(foregroundColor = "#E2EFDA", backgroundColor="#E2EFDA")
cs11 <- CellStyle(wb) + Alignment(h = "ALIGN_LEFT", v="VERTICAL_CENTER") + Fill(foregroundColor = "#E2EFDA", backgroundColor="#E2EFDA")
cs2 <- CellStyle(wb) + Alignment(h = "ALIGN_CENTER", v="VERTICAL_CENTER") + Fill(foregroundColor = "#D9E1F2", backgroundColor="#D9E1F2")
cs3 <- CellStyle(wb) + Alignment(h = "ALIGN_CENTER", v="VERTICAL_CENTER") + Fill(foregroundColor = "#FFF2CC", backgroundColor="#FFF2CC")
cs4 <- CellStyle(wb) + Alignment(h = "ALIGN_CENTER", v="VERTICAL_CENTER") + Fill(foregroundColor = "#D6DCE4", backgroundColor="#D6DCE4")

cb2 <- CellBlock(sheet2, 2, 1, samples + 5, weekNum*2*5 + 3)
cs <- CellStyle(wb) + Alignment(h = "ALIGN_CENTER", v="VERTICAL_CENTER") 

### 基本信息
# 项目信息
addMergedRegion(sheet1, 2, 2, 1, 12)
CB.setColData(cb1, "项目信息", colIndex = 1, rowOffset = 0, colStyle = cs1)
CB.setBorder(cb1, Border(color = "black",position=c("BOTTOM")), 1, 1:12)
CB.setColData(cb1, "项目编号：", colIndex = 1, rowOffset = 1, colStyle = cs11)
addMergedRegion(sheet1, 3, 3, 2, 12)
CB.setColData(cb1, "", colIndex = 2, rowOffset = 1, colStyle = cs11)
CB.setColData(cb1, "标题：", colIndex = 1, rowOffset = 2, colStyle = cs11)
addMergedRegion(sheet1, 4, 4, 2, 12)
CB.setColData(cb1, "", colIndex = 2, rowOffset = 2, colStyle = cs11)
CB.setColData(cb1, "研究目的：", colIndex = 1, rowOffset = 3, colStyle = cs11)
addMergedRegion(sheet1, 5, 5, 2, 12)
CB.setColData(cb1, "", colIndex = 2, rowOffset = 3, colStyle = cs11)

# 参与人员
addMergedRegion(sheet1, 7, 7, 1, 12)
CB.setColData(cb1, "参与人员", colIndex = 1, rowOffset = 5, colStyle = cs2)
addMergedRegion(sheet1, 8, 8, 1, 3)
CB.setColData(cb1, "职位", colIndex = 1, rowOffset = 6, colStyle = cs2)
CB.setBorder(cb1, Border(color = "black",position=c("TOP","BOTTOM")), 7, 1:3)
addMergedRegion(sheet1, 8, 8, 4, 7)
CB.setColData(cb1, "名称", colIndex = 4, rowOffset = 6, colStyle = cs2)
CB.setBorder(cb1, Border(color = "black",position=c("TOP","BOTTOM")), 7, 4:7)
addMergedRegion(sheet1, 8, 8, 8, 12)
CB.setColData(cb1, "联系方式", colIndex = 8, rowOffset = 6, colStyle = cs2)
CB.setBorder(cb1, Border(color = "black",position=c("TOP","BOTTOM")), 7, 8:12)

addMergedRegion(sheet1, 9, 9, 1, 3)
CB.setColData(cb1, "项目负责人", colIndex = 1, rowOffset = 7, colStyle = cs2)
addMergedRegion(sheet1, 9, 9, 4, 7)
CB.setColData(cb1, "", colIndex = 4, rowOffset = 7, colStyle = cs2)
addMergedRegion(sheet1, 9, 9, 8, 12)
CB.setColData(cb1, "", colIndex = 8, rowOffset = 7, colStyle = cs2)

addMergedRegion(sheet1, 10, 10, 1, 3)
CB.setColData(cb1, "技术负责人", colIndex = 1, rowOffset = 8, colStyle = cs2)
addMergedRegion(sheet1, 10, 10, 4, 7)
CB.setColData(cb1, "", colIndex = 4, rowOffset = 8, colStyle = cs2)
addMergedRegion(sheet1, 10, 10, 8, 12)
CB.setColData(cb1, "", colIndex = 8, rowOffset = 8, colStyle = cs2)

addMergedRegion(sheet1, 11, 11, 1, 3)
CB.setColData(cb1, "实验员", colIndex = 1, rowOffset = 9, colStyle = cs2)
addMergedRegion(sheet1, 11, 11, 4, 7)
CB.setColData(cb1, "", colIndex = 4, rowOffset = 9, colStyle = cs2)
addMergedRegion(sheet1, 11, 11, 8, 12)
CB.setColData(cb1, "", colIndex = 8, rowOffset = 9, colStyle = cs2)

addMergedRegion(sheet1, 12, 12, 1, 3)
CB.setColData(cb1, "实验员", colIndex = 1, rowOffset = 10, colStyle = cs2)
addMergedRegion(sheet1, 12, 12, 4, 7)
CB.setColData(cb1, "", colIndex = 4, rowOffset = 10, colStyle = cs2)
addMergedRegion(sheet1, 12, 12, 8, 12)
CB.setColData(cb1, "", colIndex = 8, rowOffset = 10, colStyle = cs2)

addMergedRegion(sheet1, 13, 13, 1, 3)
CB.setColData(cb1, "实验员", colIndex = 1, rowOffset = 11, colStyle = cs2)
addMergedRegion(sheet1, 13, 13, 4, 7)
CB.setColData(cb1,"", colIndex = 4, rowOffset = 11, colStyle = cs2)
addMergedRegion(sheet1, 13, 13, 8, 12)
CB.setColData(cb1,"", colIndex = 8, rowOffset = 11, colStyle = cs2)

CB.setBorder(cb1, Border(color = "black",position=c("RIGHT")), 7:12, 3)
CB.setBorder(cb1, Border(color = "black",position=c("RIGHT")), 7:12, 7)
CB.setBorder(cb1, Border(color = "black",position=c("BOTTOM")), 12, 1:12)

# 分组及给药情况
addMergedRegion(sheet1, 15, 15, 1, 12)
CB.setColData(cb1, "分组及给药情况", colIndex = 1, rowOffset = 13, colStyle = cs3)
addMergedRegion(sheet1, 16, 16, 1, 3)
CB.setColData(cb1, "组别", colIndex = 1, rowOffset = 14, colStyle = cs3)
CB.setBorder(cb1, Border(color = "black",position=c("TOP","BOTTOM")), 15, 1:3)
addMergedRegion(sheet1, 16, 16, 4, 12)
CB.setColData(cb1, "给药情况", colIndex = 4, rowOffset = 14, colStyle = cs3)
CB.setBorder(cb1, Border(color = "black",position=c("TOP","BOTTOM")), 15, 4:12)

for(n in 1:groupNum){
	addMergedRegion(sheet1, 16+n, 16+n, 1, 3)
	CB.setColData(cb1, "", colIndex = 1, rowOffset = 14+n, colStyle = cs3)
	addMergedRegion(sheet1, 16+n, 16+n, 4, 12)
	CB.setColData(cb1, "", colIndex = 4, rowOffset = 14+n, colStyle = cs3)
}
CB.setBorder(cb1, Border(color = "black",position=c("RIGHT")), 15:(15+n), 3)
CB.setBorder(cb1, Border(color = "black",position=c("BOTTOM")), 15+n, 1:12)


# 个体数据
for(i in 1:(weekNum*2)){
	#print(i)
	addMergedRegion(sheet2, 2, 2, (i-1)*5 + 2, i*5 + 1)
	addMergedRegion(sheet2, 3, 3, (i-1)*5 + 2, i*5 + 1)
	CB.setColData(cb2, "日期：", (i-1)*5 + 2, colStyle = cs)
	CB.setColData(cb2, "治疗天数：", rowOffset = 1, (i-1)*5 + 2, colStyle = cs)
	CB.setColData(cb2, "动物编号", rowOffset = 2, (i-1)*5 + 2, colStyle = cs)
	CB.setColData(cb2, "体重(g)", rowOffset = 2, (i-1)*5 + 3, colStyle = cs)
	CB.setColData(cb2, "长径(mm)", rowOffset = 2, (i-1)*5 + 4, colStyle = cs)
	CB.setColData(cb2, "短径(mm)", rowOffset = 2, (i-1)*5 + 5, colStyle = cs)
	CB.setColData(cb2, "TV(mm3)", rowOffset = 2, (i-1)*5 + 6, colStyle = cs)
}
border = Border(color="black", position=c("TOP", "BOTTOM", "LEFT", "RIGHT"))
CB.setBorder(cb2, border, 4: (3+samples), 1)
for(k in 1:(samples + 3)){
	CB.setBorder(cb2, border, k, 2:(weekNum*2*5 + 1))
}


for(j in 1:groupNum){
	#print(j)
	addMergedRegion(sheet2, 5 + (j-1)*repNum, j*repNum + 4, 1, 1)
	CB.setRowData(cb2, paste0("G", j), (j-1)*repNum + 4, rowStyle = cs)
}

saveWorkbook(wb, "newRecord.xlsx")

