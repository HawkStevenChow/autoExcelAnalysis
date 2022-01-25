
options("repos" = c(CRAN="https://mirror.lzu.edu.cn/CRAN/"))
if(!require(openxlsx)){install.packages("openxlsx")}
if(!require(dplyr)){install.packages("dplyr")}
if(!require(stringr)){install.packages("stringr")}
if(!require(extrafont)){install.packages("extrafont")}
if(!require(Cairo)){install.packages("Cairo")}
if(!require(ggplot2)){install.packages("ggplot2")}
if(!require(plyr)){install.packages("plyr")}
require(extrafont)
require(openxlsx)
require(plyr)

#########################
###### Input file #######
#########################
# 输入实验记录完成的EXCEL文件名称
InputFile = "recorded23.xlsx"


#########################
####### Functions #######
#########################

colList = c("black", "red", "forestgreen", "blue", "darkgoldenrod3", "darkmagenta", "olivedrab", "navyblue", "orange", "gray50",
			"deeppink", "darkslategrey", "dodgerblue3", "plum2", "springgreen")
shapeList = c(16, 17,18, 15)

removeRowsAllNa  <- function(x){x[apply(x, 1, function(y) any(!is.na(y))),]}
removeColsAllNa  <- function(x){x[, apply(x, 2, function(y) any(!is.na(y)))]}

extractDataInfo <- function(FromList = NULL, nameCols = NULL, Formula = NULL, rowNameChange = TRUE){
	FinalData <- NULL
	if(is.list(FromList)){
		if(length(nameCols) ==1){
			tmpNames <- NULL
			for(r in 1:length(FromList)){
				if(is.null(FinalData)){
					if(!is.null(FromList[[r]][[nameCols]])){
						FinalData = FromList[[r]][[nameCols]]
						if(is.null(tmpNames)){
							tmpNames = names(FromList)[r]
						}else{
							tmpNames = c(tmpNames, names(FromList)[r])
						}
					}
				}else{
					if(!is.null(FromList[[r]][[nameCols]])){
						FinalData = rbind(FinalData, FromList[[r]][[nameCols]])
						tmpNames = c(tmpNames, names(FromList)[r])
					}
				}
			}
			if(rowNameChange){
				rownames(FinalData) = tmpNames
			}
			return(FinalData)
		}else if(length(nameCols) > 1){
			#print(nameCols)
			if(length(Formula) == 1){
				#print(as.character(Formula))
				tmpNames <- NULL
				for(r in 1:length(FromList)){
					tmpLine <- NULL
					for(s in 1:length(FromList[[1]][[nameCols[1]]])){
						tmpValue <- NULL
						tmpValue = FromList[[r]][[nameCols[1]]][s]
						if(!is.na(tmpValue)){
							for(t in 2:length(nameCols)){
								if(!is.na(FromList[[r]][[nameCols[t]]][s])){
									tmpValue = paste0(tmpValue, Formula, FromList[[r]][[nameCols[t]]][s])
								}else{
									tmpValue = "NA"
								}
							}
						}else{
							tmpValue = "NA"
						}
						if(is.null(tmpLine)){
							tmpLine = tmpValue
						}else{
							tmpLine = c(tmpLine, tmpValue)
						}
					}
					if(is.null(tmpNames)){
						tmpNames = names(FromList)[r]
					}else{
						tmpNames = c(tmpNames, names(FromList)[r])
					}
					if(is.null(FinalData)){
						FinalData = tmpLine
					}else{
						FinalData = rbind(FinalData, tmpLine)
					}
				}
				if(rowNameChange){
					rownames(FinalData) = tmpNames
				}
				colnames(FinalData) = names(FromList[[1]][[nameCols[1]]])
				return(FinalData)
			}else{
				stop("\'Formula\' has no efficient value!")
			}
		}else{
			stop("\'nameCols\' has no values!")
		}
	}else{
		stop("\'FromList\' is not list!")
	}
	
}

MergeMeanSEM <- function(meanData = NULL, semData = NULL){
	newData <- NULL
	for(i in 1:nrow(meanData)){
		for(j in 1:ncol(meanData)){
			tmpLine = c(rownames(meanData)[i], colnames(meanData)[j], meanData[i,j], semData[i,j])
			if(is.null(newData)){
				newData = tmpLine
			}else{
				newData = rbind(newData,tmpLine)
			}			
		}
	}
	colnames(newData) = c("group","day","Mean","SEM")
	rownames(newData) = seq(1, nrow(newData))
	newData2 = data.frame(group = newData[,'group'], day = as.numeric(newData[,'day']), 
					means = as.numeric(newData[,'Mean']), SEM = as.numeric(newData[,'SEM']))
	newData2$group = as.character(newData2$group)
	
	return(newData2)
}

# 为每个组别添加颜色和形状，并设置为factor
addColShape <- function(dataFrame = NULL){
	newList <- list()
	colorLevels = c()
	shapeLevels = c()
	for(i in 1:length(unique(dataFrame$group))){
		tmpGroupInd = unique(dataFrame$group)[i]
		# 为每个组别设置颜色
		if(i <= length(colList)){
			dataFrame$colorType[which(dataFrame$group == tmpGroupInd)] = colList[i]
			colorLevels = c(colorLevels, colList[i])
		}else{
			newInd = i - length(colList)
			dataFrame$colorType[which(dataFrame$group == tmpGroupInd)] = colList[newInd]
			colorLevels = c(colorLevels, colList[newInd])
		}
		# 为每个组别设置形状
		if(i <= length(shapeList)){
			dataFrame$shapeType[which(dataFrame$group == tmpGroupInd)] = shapeList[i]
			shapeLevels = c(shapeLevels, shapeList[i])
		}else{
			newInd2 = i - length(shapeList)
			dataFrame$shapeType[which(dataFrame$group == tmpGroupInd)] = shapeList[newInd2]
			shapeLevels = c(shapeLevels, shapeList[newInd2])
		}
	}
	newList[['data']] = dataFrame
	newList[['color']] = colorLevels
	newList[['shape']] = shapeLevels
	return(newList)
}

# 将GroupList中每个组别的 Treatment 添加到 TCperTVdata
AddTreat2DataFrameV1 <- function(dataF = NULL, treatInfo = NULL){
	for(i in 1:length(rownames(dataF))){
		GroupID = rownames(dataF)[i]
		rownames(dataF)[which(rownames(dataF) == GroupID)] = paste(GroupID, paste(treatInfo[[GroupID]], collapse=","), sep=", ")
	}
	return(dataF)
}

AddTreat2DataFrameV2 <- function(dataF = NULL, treatInfo = NULL){
	for(i in 1:length(unique(dataF$group))){
		GroupID = unique(dataF$group)[i]
		dataF$group[which(dataF$group == GroupID)] = paste(GroupID, paste(treatInfo[[GroupID]], collapse=","), sep=", ")
	}
	return(dataF)
}

# plot of mean + SEM
PlotMeanSEM <- function(data = NULL, maintitle = "", ystart = NULL, xlab = "", ylab = "", colS = NULL, shapeS = NULL, picName = ""){
	if(min(data$means, na.rm=T) - min(data$SEM, na.rm=T) >= 0){
		if(is.null(ystart)){
			yStart = 0
		}
	}else{
		if(is.null(ystart)){
			yStart = (min(data$means, na.rm=T) - min(data$SEM, na.rm=T))*1.2
		}
	}
	
	p <- ggplot(data, aes(x=day, y=means)) + 
		geom_errorbar(aes(ymin=means-SEM, ymax=means+SEM, color = str_wrap(group,60)), width=.3, size=0.2) +
		geom_line(aes(group = str_wrap(group,60), color = str_wrap(group,60))) +
		geom_point(aes(color = str_wrap(group,60), shape = str_wrap(group,60)), size = 2) + 
		scale_color_manual(values=colS) +
		scale_shape_manual(values=shapeS) +
		scale_x_continuous(expand = c(0, 0), n.breaks = 8) +  
		scale_y_continuous(expand = c(0, 0), limits=c(yStart,(max(data$means, na.rm=T) + max(data$SEM, na.rm=T))*1.2), n.breaks = 10) +  
		labs(title=maintitle,x=xlab, y = ylab) +
		theme_classic() +
		theme(plot.title = element_text(hjust = 0.5, face = "bold"), plot.background = element_rect(color = "black", size = 0.2),
				legend.box.margin=margin(c(0,0,0,170)),
				legend.title=element_blank(), legend.key.height = unit(1.2,"cm"),
				text = element_text(family = "Sans"), legend.text = element_text(size = 8),
				axis.line = element_line(size = .2), panel.grid.major.y = element_line(colour = "black", size = 0.1))

	Cairo(width = 5000, height = 1200, file=picName, dpi=400)
	plot(p)
	dev.off()	
}

singlePlot <- function(dataset = NULL, maintitle = "", ytitle = "", picfile = "", legendMargin = TRUE, zeroStart = TRUE, yStart = 0){
	tmpdata <- NULL
	for(ii in 1:nrow(dataset)){
		for(jj in 1:ncol(dataset)){
			tmpLine = c(rownames(dataset)[ii], colnames(dataset)[jj], dataset[ii,jj])
			if(is.null(tmpdata)){
				tmpdata = tmpLine
			}else{
				tmpdata = rbind(tmpdata,tmpLine)
			}
		}
	}
	colnames(tmpdata) = c("group","day","volume")
	rownames(tmpdata) = seq(1, nrow(tmpdata))	
	finaldata = data.frame(group = tmpdata[,'group'], day = as.numeric(tmpdata[,'day']), 
					volume = as.numeric(tmpdata[,'volume']))
	# 为每个组别添加颜色和形状，并设置为factor
	colorTmpLevels = c()
	shapeTmpLevels = c()
	for(cs in 1:length(unique(finaldata$group))){
		tmpGroupInd = unique(finaldata$group)[cs]
		# 为每个组别设置颜色
		if(cs <= length(colList)){
			finaldata$colorType[which(finaldata$group == tmpGroupInd)] = colList[cs]
			colorTmpLevels = c(colorTmpLevels, colList[cs])
		}else{
			newInd = cs - length(colList)
			finaldata$colorType[which(finaldata$group == tmpGroupInd)] = colList[newInd]
			colorTmpLevels = c(colorTmpLevels, colList[newInd])
		}
		# 为每个组别设置形状
		if(cs <= length(shapeList)){
			finaldata$shapeType[which(finaldata$group == tmpGroupInd)] = shapeList[cs]
			shapeTmpLevels = c(shapeTmpLevels, shapeList[cs])
		}else{
			newInd2 = cs - length(shapeList)
			finaldata$shapeType[which(finaldata$group == tmpGroupInd)] = shapeList[newInd2]
			shapeTmpLevels = c(shapeTmpLevels, shapeList[newInd2])
		}
	}

	if(zeroStart){
		yStart = 0
	}else{
		yMin = min(finaldata$volume, na.rm=T)*1.2
		if(yMin > -10){
			yStart = -10
		}else if(yMin > -20){
			yStart = -20
		}else if(yMin > -30){
			yStart = -30
		}else if(yMin > -40){
			yStart = -40
		}
	}
	if(legendMargin){
		P <- ggplot(finaldata, aes(x=day, y=volume)) + 
				geom_line(aes(group = str_wrap(group,60), color = str_wrap(group,60)))+
				geom_point(aes(color = str_wrap(group,60), shape = str_wrap(group,60)), size = 2)+
				scale_color_manual(values=colorTmpLevels) +
				scale_shape_manual(values=shapeTmpLevels) +
				scale_x_continuous(expand = c(0, 0), limits=c(0,max(finaldata$day, na.rm=T)*1.1), n.breaks = 8) + 
				scale_y_continuous(expand = c(0, 0), limits=c(yStart,max(finaldata$volume, na.rm=T)*1.2), n.breaks = 10) +  
				labs(title=maintitle,x="Treatment Days", y = maintitle) +
				theme_classic() +
				theme(plot.title = element_text(hjust = 0.5, face = "bold"), plot.background = element_rect(color = "black", size = 0.2),
					legend.box.margin=margin(c(0,0,0,170)),
					legend.title=element_blank(), legend.key.height = unit(1.2,"cm"),
					text = element_text(family = "Sans"), legend.text = element_text(size = 8),
					axis.line = element_line(size = .2), panel.grid.major.y = element_line(colour = "black", size = 0.1))
		Cairo(width = 5000, height = 1200, file=picfile, dpi=400)
		plot(P)
		dev.off()
	}else{
		P <- ggplot(finaldata, aes(x=day, y=volume)) + 
				geom_line(aes(group = str_wrap(group,60), color = str_wrap(group,60)))+
				geom_point(aes(color = str_wrap(group,60), shape = str_wrap(group,60)), size = 2)+
				scale_color_manual(values=colorTmpLevels) +
				scale_shape_manual(values=shapeTmpLevels) +
				scale_x_continuous(expand = c(0, 0), limits=c(0,max(finaldata$day, na.rm=T)*1.1), n.breaks = 8) + 
				scale_y_continuous(expand = c(0, 0), limits=c(yStart,max(finaldata$volume, na.rm=T)*1.2), n.breaks = 10) +  
				labs(title=maintitle,x="Treatment Days", y = ytitle) +
				theme_classic() +
				theme(plot.title = element_text(hjust = 0.5, face = "bold"), plot.background = element_rect(color = "black", size = 0.2),
					legend.title=element_blank(), legend.key.height = unit(1,"cm"),
					text = element_text(family = "Sans"), legend.text = element_text(size = 8),
					axis.line = element_line(size = .2), panel.grid.major.y = element_line(colour = "black", size = 0.1))
		Cairo(width = 3000, height = 1200, file=picfile, dpi=400)
		plot(P)
		dev.off()	
	}
}


#########################
####### Functions #######
#########################


###################################
###################################
####### read recorded file ########
###################################
###################################

########################
### read basic infos ###
########################
# 读取记录结果的第一个工作簿
BasicData = openxlsx::read.xlsx(InputFile, sheet = 1, startRow = 1, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)
print("Reading the first sheet of EXCEL file......\n")
# 获取每个板块的起始行
GIindex = which(BasicData$X1 == "General infos")
PEindex = which(BasicData$X1 == "Personel")
ATindex = which(BasicData$X1 == "Treatments")
# 批量定义变量
projectID <- NULL;projectTitle <- NULL;Object <- NULL
StudyDirector <- NULL;TechGroupLeader <- NULL;LabTechList <- NULL
GroupList <- list()
GroupTreatData <- NULL
# 根据每个板块的起始行间隔，逐个读取板块内容
if(length(GIindex) !=0 & length(PEindex) != 0 & length(ATindex) !=0){
	# 读取 General infos 板块内容
	for(gi in (GIindex + 1):(PEindex - 2)){
		if(grepl("Study Number", BasicData$X1[gi])){
			projectID = BasicData$X3[gi]
		}else if(grepl("Title", BasicData$X1[gi])){
			projectTitle = BasicData$X3[gi]
		}else if(grepl("Objective", BasicData$X1[gi])){
			Object = BasicData$X3[gi]
		}
	}
	# 读取 Personel 板块内容
	LabNum = 0
	for(ei in (PEindex + 1):(ATindex - 2)){
		if(grepl("Study Director", BasicData$X1[ei])){
			StudyDirector = c(as.character(BasicData$X4[ei]), as.character(BasicData$X9[ei]))
		}else if(grepl("Technician Group Leader", BasicData$X1[ei])){
			TechGroupLeader = c(as.character(BasicData$X4[ei]), as.character(BasicData$X9[ei]))
		}else if(grepl("Lab Technician", BasicData$X1[ei])){
			if(!is.na(BasicData$X4[ei])){
				LabNum = LabNum + 1
				LabTechList[[LabNum]] = c(as.character(BasicData$X4[ei]), as.character(BasicData$X9[ei]))
			}
		}
	}
	# 读取 Articles 板块内容
	groupName = c()
	for(ai in (ATindex + 2):nrow(BasicData)){
		if(!BasicData$X1[ai] %in% groupName){
			groupName = c(groupName, BasicData$X1[ai])
			GroupList[[BasicData$X1[ai]]] = BasicData$X4[ai]
		}else{
			#GroupList[[BasicData$X1[ai]]] = paste(GroupList[[BasicData$X1[ai]]], BasicData$X4[ai], sep=";")
			preTmpList = GroupList[[BasicData$X1[ai]]]
			curTmpList = c(preTmpList, BasicData$X4[ai])
			GroupList[[BasicData$X1[ai]]] <- NULL
			GroupList[[BasicData$X1[ai]]] = curTmpList
		}
	}
}else{
	if(length(GIindex) ==0){
		print("There is no \'General infos\' in the first sheet!")
	}
	if(length(PEindex) ==0){
		print("There is no \'Personel\' in the first sheet!")
	}
	if(length(ATindex) ==0){
		print("There is no \'Articles\' in the first sheet!")
	}
}
TreatList <- c()
for(gl in 1:length(GroupList)){
	TreatList <- c(TreatList, paste(names(GroupList)[gl], paste(GroupList[[gl]], collapse = ", "), sep=", "))
}
GroupTreatData <- data.frame(Group = names(GroupList), Caption = TreatList, stringsAsFactors = FALSE)

##########################
### read Dosing record ###
##########################
# 读取记录结果的第二个工作簿
DoseRecord = openxlsx::read.xlsx(InputFile, sheet = 2, startRow = 1, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)

## Date --- Days ##
# 读取 有效记录天数Days
dayName = c()
for(dd in 6:ncol(DoseRecord)){
	day = DoseRecord[7,dd]
	dayfmt = gsub("Days:", "",day)
	if(dayfmt != ""){
		if(!dayfmt %in% dayName){
			dayName = c(dayName, dayfmt)
		}else{
			stop("Day number in sheet \'DosingRecords\' has repeat record!")
		}
	}
}
# 获取 有效记录日期Date
dateName = c()
for(tt in 6:ncol(DoseRecord)){
	dated = DoseRecord[6,tt]
	datedfmt = gsub("Date:", "",dated)
	if(datedfmt != ""){
		if(!datedfmt %in% dateName){
			dateName = c(dateName, datedfmt)
		}else{
			stop("Date number in sheet \'DosingRecords\' has repeat record!")
		}
	}
}
# 根据已获取 天数Days 和 日期Date的结果进行匹配，
# 1）当天数和日期数 匹配不上，停止程序并报错；
# 2）当天数和日期数匹配，将日期和天数一一对应，并存放在新列表中；
days2Date = list()
if(length(dayName) != length(dateName)){
	stop("Date record is different with days record in sheet 3!")
}else{
	days2Date = dateName
	names(days2Date) = dayName
}

## Group --- Samples ##
# 读取 有效组别Groups和组别中的样本Samples
# 根据已读取工作簿的第1列的第7行起，根据表格内容是否为NA，来获取每个组别的起始行
# 根据每个组别的起始行间隔，逐个读取单个组别中的样本号
# 当某个组别有多个处理条件时，需要确认样本数量和名称是否一致
groupIndex <- which(!is.na(DoseRecord$X1[8:nrow(DoseRecord)]))
groupName2 = c()
SampleList <- list()
for(gs in 1:length(groupIndex)){
	if(gs != length(groupIndex)){
		if(!DoseRecord$X1[groupIndex[gs] + 7] %in% groupName2){
		# 当 组别 未被记录时
			groupName2 = c(groupName2, DoseRecord$X1[groupIndex[gs] + 7])
			sampleArray <- c()
			# 根据与后一个组别的起始行间隔，读取该组别的样本名称
			for(i in (groupIndex[gs] + 7):(groupIndex[gs+1] + 6)){
				if(!is.na(DoseRecord$X5[i])){
					sampleArray <- c(sampleArray, DoseRecord$X5[i])
				}
			}
			SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]] = sampleArray
		}else{
		# 当组别已有记录时
			# 确认该组别的不同处理条件中的样本数量和名称是否一致，
			# 如果否，停止程序并报错
			sampleArray <- c()
			for(i in (groupIndex[gs] + 7):(groupIndex[gs+1] + 6)){
				if(!is.na(DoseRecord$X5[i])){
					sampleArray <- c(sampleArray, DoseRecord$X5[i])
				}
			}
			if(!identical(SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]], sampleArray)){
				print(SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]])
				print(sampleArray)
				stop(paste0(DoseRecord$X1[groupIndex[gs] + 7], "with different article have different samples!"))
			}
		}
	}else{
	# 当最后一个组别时，
		# 根据‘与后一个组别的起始行间隔，读取该组别的样本名称’的方法不再适用
		# 根据与已读取工作簿内容最后一行的间隔，读取最后一个组别的样本名称
		if(!DoseRecord$X1[groupIndex[gs] + 7] %in% groupName2){
		# 当 组别 未被记录时
			groupName2 = c(groupName2, DoseRecord$X1[groupIndex[gs] + 7])
			sampleArray <- c()
			# 根据与已读取工作簿内容最后一行的间隔，读取最后一个组别的样本名称
			for(i in (groupIndex[gs] + 7):length(DoseRecord$X1)){
				if(!is.na(DoseRecord$X5[i])){
					sampleArray <- c(sampleArray, DoseRecord$X5[i])
				}
			}
			SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]] = sampleArray
		}else{
		# 当组别已有记录时
			# 确认该组别的不同处理条件中的样本数量和名称是否一致，
			# 如果否，停止程序并报错
			sampleArray <- c()
			for(i in (groupIndex[gs] + 7):length(DoseRecord$X1)){
				if(!is.na(DoseRecord$X5[i])){
					sampleArray <- c(sampleArray, DoseRecord$X5[i])
				}
			}
			if(!identical(SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]], sampleArray)){
				print(SampleList[[DoseRecord$X1[groupIndex[gs] + 7]]])
				print(sampleArray)
				stop(paste0(DoseRecord$X1[groupIndex[gs] + 7], "with different article have different samples!"))
			}
		}
	}
}

## samples --- Days Record ##
# 读取 每个样本的有效记录天数
sampleIndex = which(!is.na(DoseRecord$X5))
sampleIndex = sampleIndex[which(sampleIndex > 7)]
# 生成以样本为行、记录天数为列的矩阵
SampleRecord <- matrix(nrow=length(sampleIndex),ncol=length(dayName))
for(di in 1:length(dayName)){
	for(si in 1:length(sampleIndex)){
		if(is.na(DoseRecord[sampleIndex[si], di + 5])){
			SampleRecord[si, di] = "N/A"
		}else{
			if(tolower(DoseRecord[sampleIndex[si], di + 5]) == "v"){
				SampleRecord[si, di] = "√"
			}else if(tolower(DoseRecord[sampleIndex[si], di + 5]) == "holiday"){
				SampleRecord[si, di] = "Holiday"
			}else{
				SampleRecord[si, di] = "N/A"
			}
		}
	}
}
SampleRecord = cbind(DoseRecord$X5[sampleIndex], SampleRecord)

##########################
#### read Sample List ####
##########################
# 读取记录结果的第三个工作簿
SampleFile = openxlsx::read.xlsx(InputFile, sheet = 3, startRow = 1, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)
# 获取 有效实体组织列
solidList = as.character(sapply(SampleFile[7,which(!is.na(SampleFile[7,]))],as.character))
# 删除“Group”和“SampleID”两个字符串
solidList = solidList[3:length(solidList)]
solidIndex = which(SampleFile[7,] %in% solidList)
# 获得每个组别在EXCEL中的起始行
groupIndex2 <- which(!is.na(SampleFile[,1]))
groupIndex2 = groupIndex2[6:length(groupIndex2)]
# 根据已读取的数据矩阵，进行基本的整理
SDlist <- list()
for(gsI in 1:length(groupIndex2)){
	groupTmpName2 <- NULL
	sampleTmpList2 <- NULL
	if(gsI != length(groupIndex2)){
		groupTmpName2 = SampleFile[groupIndex2[gsI], 1]
		sampleTmpList2 = (groupIndex2[gsI]:(groupIndex2[gsI+1] - 1))[which(!is.na(SampleFile[groupIndex2[gsI]:(groupIndex2[gsI+1] - 1),2]))]	
	}else{
		groupTmpName2 = SampleFile[groupIndex2[gsI], 1]
		sampleTmpList2 = (groupIndex2[gsI]:length(SampleFile$X1))[which(!is.na(SampleFile[groupIndex2[gsI]:length(SampleFile$X1),2]))]
	}
	SDlist[[groupTmpName2]][['samples']] = as.character(SampleFile[sampleTmpList2, 2])
	tmpSampleLiData <- SampleFile[sampleTmpList2, solidIndex]
	rownames(tmpSampleLiData) = as.character(SampleFile[sampleTmpList2, 2])
	colnames(tmpSampleLiData) = solidList
	SDlist[[groupTmpName2]][['data']] = tmpSampleLiData	
}

##########################
#### read Date record ####
##########################
# 读取记录结果的第三个工作簿
DateRecord = openxlsx::read.xlsx(InputFile, sheet = 4, startRow = 1, colNames = FALSE, skipEmptyCols = FALSE, skipEmptyRows = FALSE)
# 获取 有效记录天数和日期
dateList = sapply(DateRecord[4,which(!is.na(DateRecord[4,]))],as.character)
dayCount = sapply(DateRecord[5,which(!is.na(DateRecord[5,]))],as.character)
recordDate <- sapply(dateList, function(x){unlist(strsplit(x, ":"))[2]})
recordDate = recordDate[which(!is.na(recordDate))]
recordDay <- sapply(dayCount, function(x){unlist(strsplit(x, ":"))[2]})
recordDay = recordDay[which(!is.na(recordDay))]
# 根据已获取 天数Days 和 日期Date的结果进行匹配，
# 1）当天数和日期数 匹配不上，停止程序并报错；
# 2）当天数和日期数匹配，将日期和天数一一对应，并存放在新列表中；
days2Date2 = list()
if(length(recordDay) != length(recordDate)){
	stop("Date record is different with days record in sheet 4!")
}else{
	days2Date2 = recordDate
	names(days2Date2) = recordDay
}
# 批量定义变量
BWdata <- list()
BWChangedata <- list()
TVdata <- list()
TVLendata <- list()
TVWiddata <- list()
BWdataFinal <- list()
BWChangedataFinal <- list()
TVdataFinal <- list()
BWmatrix <- NULL
BWChangematrix <- NULL
TVmatrix <- NULL
library(dplyr)
# 获得每个组别在EXCEL中的起始行
groupIndex2 <- which(!is.na(DateRecord[,1]))
groupIndex2 = groupIndex2[4:length(groupIndex2)]
# 根据列名获得重量和体积的index(只收录有天数记录的index)
BWindex <- which(DateRecord[6,] == "BodyWeight(g)")[1:length(recordDay)]
TVindex <- which(DateRecord[6,] == "TV(mm3)")[1:length(recordDay)]
LENindex <- which(DateRecord[6,] == "Length(mm)")[1:length(recordDay)]
WIDindex <- which(DateRecord[6,] == "Width(mm)")[1:length(recordDay)]
# 根据已读取的数据矩阵，进行基本的计算，并保存结果
for(gt in 1:length(groupIndex2)){
	if(gt != length(groupIndex2)){
		groupTmpName = DateRecord[groupIndex2[gt], 1]
		sampleTmpList = (groupIndex2[gt]:(groupIndex2[gt+1] - 1))[which(!is.na(DateRecord[groupIndex2[gt]:(groupIndex2[gt+1] - 1),2]))]
		#############################################################
		## get body weight of all samples in this group into 'BWdata'
		#############################################################
		# raw data of BodyWeight
		tmpBwData <- DateRecord[sampleTmpList, BWindex]
		tmpBwData = apply(tmpBwData, 2, function(x){as.numeric(x)})
		rownames(tmpBwData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpBwData) = as.character(recordDay)
		BWdata[[groupTmpName]][['data']] = tmpBwData
		tmpBwDataPer = apply(tmpBwData, 2, function(x){round((x - tmpBwData[,1])/tmpBwData[,1] * 100, 2)})
		# calculate result of BodyWeight
		BWdata[[groupTmpName]][['Mean']] = apply(tmpBwData, 2, function(x){round(mean(x),2)})
		BWdata[[groupTmpName]][['StdDev']] = apply(tmpBwData, 2, function(x){round(sd(x),2)})
		BWdata[[groupTmpName]][['Count']] = apply(tmpBwData, 2, function(x){length(x[which(!is.na(x))])})
		BWdata[[groupTmpName]][['SEM']] = apply(tmpBwData, 2, function(x){round(sd(x)/sqrt(length(x)),2)})
		BWdata[[groupTmpName]][['Δ%Change']] = apply(tmpBwData, 2, function(x){round((mean(x) - mean(tmpBwData[,1]))/mean(tmpBwData[,1]) * 100, 2)})
		BWdata[[groupTmpName]][['Δ%Change']][1] = NA	
		# merge one group data and calculate result into one list
		BWdataFinal[[groupTmpName]] = do.call(rbind, BWdata[[groupTmpName]])
		if(is.null(BWmatrix)){
			BWmatrix <- do.call(rbind, BWdata[[groupTmpName]])
		}else{
			BWmatrix <- rbind(BWmatrix, do.call(rbind, BWdata[[groupTmpName]]))
		}	
		
		# get 'BW change %' data
		tmpBwDataPer = apply(tmpBwData, 2, function(x){round((x - tmpBwData[,1])/tmpBwData[,1] * 100, 2)})
		BWChangedata[[groupTmpName]][['data']] = tmpBwDataPer
		# calculate result of 'BW change %' data
		BWChangedata[[groupTmpName]][['%Ch Mean']] = apply(tmpBwDataPer, 2, function(x){round(mean(x),2)})
		BWChangedata[[groupTmpName]][['%Ch StdDev']] = apply(tmpBwDataPer, 2, function(x){round(sd(x),2)})
		BWChangedata[[groupTmpName]][['Count']] = apply(tmpBwDataPer, 2, function(x){length(x[which(!is.na(x))])})
		BWChangedata[[groupTmpName]][['%Ch SEM']] = apply(tmpBwDataPer, 2, function(x){round(sd(x)/sqrt(length(x)),2)})		
		# BWChangedata[[groupTmpName]][['DeltaChangePer']] = apply(tmpBwDataPer, 2, function(x){return(NA)})	
		# merge one group 'BW change %' data and calculate result into one list
		BWChangedataFinal[[groupTmpName]] = do.call(rbind, BWChangedata[[groupTmpName]])
		if(is.null(BWChangematrix)){
			BWChangematrix <- do.call(rbind, BWChangedata[[groupTmpName]])
		}else{
			BWChangematrix <- rbind(BWChangematrix, do.call(rbind, BWChangedata[[groupTmpName]]))
		}
		
		###############################################################
		## get tumor volume of all samples in this group into 'TVdata'
		###############################################################
		# raw data of Tumor Volume
		tmpTvData <- DateRecord[sampleTmpList, TVindex]
		tmpTvData = apply(tmpTvData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvData) = as.character(recordDay)
		TVdata[[groupTmpName]][['data']] = tmpTvData
		# raw data of length of Tumor Volume
		tmpTvLenData <- DateRecord[sampleTmpList, LENindex]
		tmpTvLenData = apply(tmpTvLenData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvLenData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvLenData) = as.character(recordDay)
		TVLendata[[groupTmpName]][['Length']] = tmpTvLenData
		# raw data of width of Tumor Volume
		tmpTvWidData <- DateRecord[sampleTmpList, WIDindex]
		tmpTvWidData = apply(tmpTvWidData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvWidData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvWidData) = as.character(recordDay)
		TVWiddata[[groupTmpName]][['Width']] = tmpTvWidData
		
		# calculate result of Tumor Volume
		TVdata[[groupTmpName]][['Mean']] = round(apply(tmpTvData, 2, function(x){round(mean(x),2)}), 2)
		TVdata[[groupTmpName]][['StdDev']] = round(apply(tmpTvData, 2, function(x){round(sd(x),2)}), 2)
		TVdata[[groupTmpName]][['Count']] = apply(tmpTvData, 2, function(x){length(x[which(!is.na(x))])})
		TVdata[[groupTmpName]][['SEM']] = round(apply(tmpTvData, 2, function(x){sd(x)/sqrt(length(x))}), 2)
		
		if(gt == 1){
			# TVdata[[groupTmpName]][['CRTV']] = sapply(TVdata[[groupTmpName]][['Mean']], function(x){round(x/TVdata[[groupTmpName]][['Mean']][1],2)})
			# TVdata[[groupTmpName]][['CRTV']][1] = NA
			# TVStatisLen = length(TVdata[[groupTmpName]][['CRTV']])
			TVdata[[groupTmpName]][['%T/C']] = rep(NA, length(recordDay))
			TVdata[[groupTmpName]][['%Inhibition']] = rep(NA, length(recordDay))
			TVdata[[groupTmpName]][['%ΔT/ΔC']] = rep(NA, length(recordDay))
			TVdata[[groupTmpName]][['%Δinhibition']] = rep(NA, length(recordDay))
			# TVdata[[groupTmpName]][['P']] = rep(NA, length(recordDay))
		}else{
			# TVdata[[groupTmpName]][['TRTV']] = sapply(TVdata[[groupTmpName]][['Mean']], function(x){round(x/TVdata[[groupTmpName]][['Mean']][1],2)})
			# TVdata[[groupTmpName]][['TRTV']][1] = NA
			# TVStatisLen = length(TVdata[[groupTmpName]][['TRTV']])
			TVdata[[groupTmpName]][['%T/C']] = sapply(seq(1,length(recordDay)), function(x){round(TVdata[[groupTmpName]][['Mean']][x] / TVdata[[1]][['Mean']][x] * 100, 2)})
			TVdata[[groupTmpName]][['%Inhibition']] = sapply(seq(1,length(recordDay)), function(x){round((TVdata[[1]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][x]) / TVdata[[1]][['Mean']][x] * 100, 2)})
			# TVdata[[groupTmpName]][['%Inhibition']][1] = NA
			TVdata[[groupTmpName]][['%ΔT/ΔC']] = sapply(seq(1,length(recordDay)), function(x){round((TVdata[[groupTmpName]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][1]) / (TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) * 100, 2)})
			TVdata[[groupTmpName]][['%ΔT/ΔC']][1] = 100.00
			TVdata[[groupTmpName]][['%Δinhibition']] = sapply(seq(1,length(recordDay)), function(x){round(((TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) - (TVdata[[groupTmpName]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][1]))/ (TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) * 100, 2)})
			TVdata[[groupTmpName]][['%Δinhibition']][1] = NA
			pvalue = c()
			for(pid in 1:length(recordDay)){
				if((length(which(!is.na(TVdata[[groupTmpName]]$data[,pid]))) >= 2) & (!length(which(is.na(TVdata[[1]]$data[,pid]))) >= 2)){
					tmpP = round(t.test(TVdata[[groupTmpName]]$data[,pid], TVdata[[1]]$data[,pid], var.equal = T)$p.value,5)
					pvalue = c(pvalue, tmpP)
				}else{
					pvalue = c(pvalue, NA)
				}
			}
			# TVdata[[groupTmpName]][['P']] = pvalue
		}
		
		TVdataFinal[[groupTmpName]] = do.call(rbind, TVdata[[groupTmpName]])		
		if(is.null(TVmatrix)){
			TVmatrix <- do.call(rbind, TVdata[[groupTmpName]])
		}else{
			TVmatrix <- rbind(TVmatrix, do.call(rbind, TVdata[[groupTmpName]]))
		}
	}else{
	# 当最后一个组别时，
		# 根据‘与后一个组别的起始行间隔，读取该组别的样本名称’的方法不再适用
		# 根据与已读取工作簿内容最后一行的间隔，读取最后一个组别的样本名称	
		groupTmpName = DateRecord[groupIndex2[gt], 1]
		sampleTmpList = (groupIndex2[gt]:length(DateRecord$X1))[which(!is.na(DateRecord[groupIndex2[gt]:length(DateRecord$X1),2]))]
		#############################################################
		## get body weight of all samples in this group into 'BWdata'
		#############################################################
		# raw data of BodyWeight
		tmpBwData <- DateRecord[sampleTmpList, BWindex]
		tmpBwData = apply(tmpBwData, 2, function(x){as.numeric(x)})
		rownames(tmpBwData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpBwData) = as.character(recordDay)
		BWdata[[groupTmpName]][['data']] = tmpBwData
		# calculate result of BodyWeight
		BWdata[[groupTmpName]][['Mean']] = apply(tmpBwData, 2, function(x){round(mean(x),2)})
		BWdata[[groupTmpName]][['StdDev']] = apply(tmpBwData, 2, function(x){round(sd(x),2)})
		BWdata[[groupTmpName]][['Count']] = apply(tmpBwData, 2, function(x){length(x[which(!is.na(x))])})
		BWdata[[groupTmpName]][['SEM']] = apply(tmpBwData, 2, function(x){round(sd(x)/sqrt(length(x)),2)})
		BWdata[[groupTmpName]][['Δ%Change']] = apply(tmpBwData, 2, function(x){round((mean(x) - mean(tmpBwData[,1]))/mean(tmpBwData[,1]) * 100, 2)})
		BWdata[[groupTmpName]][['Δ%Change']][1] = NA	
		# merge one group data and calculate result into one list
		BWdataFinal[[groupTmpName]] = do.call(rbind, BWdata[[groupTmpName]])
		if(is.null(BWmatrix)){
			BWmatrix <- do.call(rbind, BWdata[[groupTmpName]])
		}else{
			BWmatrix <- rbind(BWmatrix, do.call(rbind, BWdata[[groupTmpName]]))
		}	
		
		# get 'BW change %' data
		tmpBwDataPer = apply(tmpBwData, 2, function(x){round((x - tmpBwData[,1])/tmpBwData[,1] * 100, 2)})
		BWChangedata[[groupTmpName]][['data']] = tmpBwDataPer
		# calculate result of 'BW change %' data
		BWChangedata[[groupTmpName]][['%Ch Mean']] = apply(tmpBwDataPer, 2, function(x){round(mean(x),2)})
		BWChangedata[[groupTmpName]][['%Ch StdDev']] = apply(tmpBwDataPer, 2, function(x){round(sd(x),2)})
		BWChangedata[[groupTmpName]][['Count']] = apply(tmpBwDataPer, 2, function(x){length(x[which(!is.na(x))])})
		BWChangedata[[groupTmpName]][['%Ch SEM']] = apply(tmpBwDataPer, 2, function(x){round(sd(x)/sqrt(length(x)),2)})		
		# BWChangedata[[groupTmpName]][['DeltaChangePer']] = apply(tmpBwDataPer, 2, function(x){return(NA)})	
		# merge one group 'BW change %' data and calculate result into one list
		BWChangedataFinal[[groupTmpName]] = do.call(rbind, BWChangedata[[groupTmpName]])
		if(is.null(BWChangematrix)){
			BWChangematrix <- do.call(rbind, BWChangedata[[groupTmpName]])
		}else{
			BWChangematrix <- rbind(BWChangematrix, do.call(rbind, BWChangedata[[groupTmpName]]))
		}
		
		###############################################################
		## get tumor volume of all samples in this group into 'TVdata'
		###############################################################
		# raw data of Tumor Volume
		tmpTvData <- DateRecord[sampleTmpList, TVindex]
		tmpTvData = apply(tmpTvData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvData) = as.character(recordDay)
		TVdata[[groupTmpName]][['data']] = tmpTvData
		# raw data of length of Tumor Volume
		tmpTvLenData <- DateRecord[sampleTmpList, LENindex]
		tmpTvLenData = apply(tmpTvLenData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvLenData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvLenData) = as.character(recordDay)
		TVLendata[[groupTmpName]][['Length']] = tmpTvLenData
		# raw data of width of Tumor Volume
		tmpTvWidData <- DateRecord[sampleTmpList, WIDindex]
		tmpTvWidData = apply(tmpTvWidData, 2, function(x){round(as.numeric(x),2)})
		rownames(tmpTvWidData) = as.character(DateRecord[sampleTmpList, 2])
		colnames(tmpTvWidData) = as.character(recordDay)
		TVWiddata[[groupTmpName]][['Width']] = tmpTvWidData
		
		# calculate result of Tumor Volume
		TVdata[[groupTmpName]][['Mean']] = round(apply(tmpTvData, 2, mean), 2)
		TVdata[[groupTmpName]][['StdDev']] = round(apply(tmpTvData, 2, sd), 2)
		TVdata[[groupTmpName]][['Count']] = apply(tmpTvData, 2, function(x){length(x[which(!is.na(x))])})
		TVdata[[groupTmpName]][['SEM']] = round(apply(tmpTvData, 2, function(x){sd(x)/sqrt(length(x))}), 2)

		# TVdata[[groupTmpName]][['TRTV']] = sapply(TVdata[[groupTmpName]][['Mean']], function(x){round(x/TVdata[[groupTmpName]][['Mean']][1],2)})
		# TVdata[[groupTmpName]][['TRTV']][1] = NA
		# TVStatisLen = length(TVdata[[groupTmpName]][['TRTV']])
		TVdata[[groupTmpName]][['%T/C']] = sapply(seq(1,length(recordDay)), function(x){round(TVdata[[groupTmpName]][['Mean']][x] / TVdata[[1]][['Mean']][x] * 100, 2)})
		TVdata[[groupTmpName]][['%Inhibition']] = sapply(seq(1,length(recordDay)), function(x){round((TVdata[[1]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][x]) / TVdata[[1]][['Mean']][x] * 100, 2)})
		# TVdata[[groupTmpName]][['%Inhibition']][1] = NA
		TVdata[[groupTmpName]][['%ΔT/ΔC']] = sapply(seq(1,length(recordDay)), function(x){round((TVdata[[groupTmpName]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][1]) / (TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) * 100, 2)})
		TVdata[[groupTmpName]][['%ΔT/ΔC']][1] = 100.00
		TVdata[[groupTmpName]][['%Δinhibition']] = sapply(seq(1,length(recordDay)), function(x){round(((TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) - (TVdata[[groupTmpName]][['Mean']][x] - TVdata[[groupTmpName]][['Mean']][1]))/ (TVdata[[1]][['Mean']][x] - TVdata[[1]][['Mean']][1]) * 100, 2)})
		TVdata[[groupTmpName]][['%Δinhibition']][1] = NA
		pvalue = c()
		for(pid in 1:length(recordDay)){
			if((length(which(!is.na(TVdata[[groupTmpName]]$data[,pid]))) >= 2) & (!length(which(is.na(TVdata[[1]]$data[,pid]))) >= 2)){
				tmpP = round(t.test(TVdata[[groupTmpName]]$data[,pid], TVdata[[1]]$data[,pid], var.equal = T)$p.value,5)
				pvalue = c(pvalue, tmpP)
			}else{
				pvalue = c(pvalue, NA)
			}
		}
		# TVdata[[groupTmpName]][['P']] = pvalue
		
		TVdataFinal[[groupTmpName]] = do.call(rbind, TVdata[[groupTmpName]])		
		if(is.null(TVmatrix)){
			TVmatrix <- do.call(rbind, TVdata[[groupTmpName]])
		}else{
			TVmatrix <- rbind(TVmatrix, do.call(rbind, TVdata[[groupTmpName]]))
		}
	}
}

# 基本的数据处理
### TV ####
TVmeanData <- extractDataInfo(FromList = TVdata, nameCols = "Mean", Formula = NULL)
TVsemData <- extractDataInfo(FromList = TVdata, nameCols = "SEM", Formula = NULL)
TVdeData <- extractDataInfo(FromList = TVdata, nameCols = c("Mean", "SEM"), Formula = "±")

### BW ####
BWmeanData <- extractDataInfo(FromList = BWdata, nameCols = "Mean", Formula = NULL)
BWsemData <- extractDataInfo(FromList = BWdata, nameCols = "SEM", Formula = NULL)
BWdeData <- extractDataInfo(FromList = BWdata, nameCols = c("Mean", "SEM"), Formula = "±")

BWmeanPerData <- extractDataInfo(FromList = BWChangedata, nameCols = "%Ch Mean", Formula = NULL)
BWsemPerData <- extractDataInfo(FromList = BWChangedata, nameCols = "%Ch SEM", Formula = NULL)
BWdePerData <- extractDataInfo(FromList = BWChangedata, nameCols = c("%Ch Mean", "%Ch SEM"), Formula = "±")

###################################
###################################
####### read recorded file ########
###################################
###################################

###################################
##### start of plots ##############
###################################

#########################################
### 'Mean ± SEM' plot of Tumor Volume ###
#########################################
# 合并体积数据的 平均值(mean) 和 标准误差(sem) 为一个数据框(dataframe)
TVde4Plot <- MergeMeanSEM(meanData = TVmeanData, semData = TVsemData)
# 为每个组别添加颜色和形状，并设置为factor
TVde4PlotList <- addColShape(dataFrame = TVde4Plot)
TVde4newPlot <- TVde4PlotList$data
colorTVLevels <- TVde4PlotList$color
shapeTVLevels <- TVde4PlotList$shape
# 将GroupList中每个组别的 Treatment 添加到 TVde4newPlot$group 
TVde4newPlot = AddTreat2DataFrameV2(dataF = TVde4newPlot, treatInfo = GroupList)
# plot 
PlotMeanSEM(data = TVde4newPlot, maintitle = "Tumor Volume Mean ± SEM", colS = colorTVLevels, shapeS = shapeTVLevels,
					xlab = "Treatment Days", ylab = expression("Tumor Volume("*mm^3*")"), picName = "TVde4Plot.png")

#########################################
###### Other plots of Tumor Volume ######
#########################################
# %T/C Tumor Volume
TCperTVdata <- extractDataInfo(FromList = TVdata, nameCols = "%T/C", Formula = NULL)
# 将GroupList中每个组别的 Treatment 添加到 TCperTVdata 
TCperTVdata = AddTreat2DataFrameV1(dataF = TCperTVdata, treatInfo = GroupList)
singlePlot(dataset = TCperTVdata, maintitle = "%T/C Tumor Volume", picfile = "TVtc4Plot.png", zeroStart = TRUE)

# %Inhibition Tumor Volume
InhPerTVdata <- extractDataInfo(FromList = TVdata, nameCols = "%Inhibition", Formula = NULL)
# 将GroupList中每个组别的 Treatment 添加到 InhPerTVdata
InhPerTVdata = AddTreat2DataFrameV1(dataF = InhPerTVdata, treatInfo = GroupList)
singlePlot(dataset = InhPerTVdata, maintitle = "%Inhibition Tumor Volume", picfile = "TVInh4Plot.png", zeroStart = FALSE)

# %ΔT/ΔC Tumor Volume
TCdeltaPerTVdata <- extractDataInfo(FromList = TVdata, nameCols = "%ΔT/ΔC", Formula = NULL)
# 将GroupList中每个组别的 Treatment 添加到 TCdeltaPerTVdata
TCdeltaPerTVdata = AddTreat2DataFrameV1(dataF = TCdeltaPerTVdata, treatInfo = GroupList)
singlePlot(dataset = TCdeltaPerTVdata, maintitle = "%ΔT/ΔC Tumor Volume", picfile = "TVtcDeltaPer4Plot.png", zeroStart = FALSE)

# %ΔInhibition Tumor Volume
InhDeltaPerTVdata <- extractDataInfo(FromList = TVdata, nameCols = "%Δinhibition", Formula = NULL)
# 将GroupList中每个组别的 Treatment 添加到 InhDeltaPerTVdata
InhDeltaPerTVdata = AddTreat2DataFrameV1(dataF = InhDeltaPerTVdata, treatInfo = GroupList)
singlePlot(dataset = InhDeltaPerTVdata, maintitle = "%ΔInhibition Tumor Volume", picfile = "TVInhDeltaPer4Plot.png", zeroStart = FALSE)

########################################
### 'Mean ± SEM' plot of Body Weight ###
########################################
# 合并体重数据的 平均值(mean) 和 标准误差(sem) 为一个数据框(dataframe)
BWde4Plot <- MergeMeanSEM(meanData = BWmeanData, semData = BWsemData)
# 为每个组别添加颜色和形状，并设置为factor
BWde4PlotList <- addColShape(dataFrame = BWde4Plot)
BWde4newPlot <- BWde4PlotList$data
colorBWLevels <- BWde4PlotList$color
shapeBWLevels <- BWde4PlotList$shape
# 将GroupList中每个组别的 Treatment 添加到 BWde4newPlot$group
BWde4newPlot = AddTreat2DataFrameV2(dataF = BWde4newPlot, treatInfo = GroupList)
# plot
PlotMeanSEM(data = BWde4newPlot, maintitle = "Body Weight Mean ± SEM", colS = colorBWLevels, shapeS = shapeBWLevels,
					xlab = "Treatment Days", ylab = "Body Weight(g)", picName = "BWde4Plot.png")

for(gb in 1:length(BWdata)){
	tmpGrName = names(BWdata)[gb]
	tmpGrData = BWdata[[gb]]$data
	singlePlot(dataset = tmpGrData, maintitle = paste("Absolute Body Weight", tmpGrName, sep=", "), ytitle = "Body Weight (g)", 
					picfile = paste0("Abs_BW_", gb, ".png"), legendMargin = FALSE, zeroStart = TRUE, yStart = 0)
}

for(gv in 1:length(TVdata)){
	tmpGrName = names(TVdata)[gv]
	tmpGrData = TVdata[[gv]]$data
	singlePlot(dataset = tmpGrData, maintitle = paste("Absolute Tumor Volume", tmpGrName, sep=", "), ytitle = expression("Tumor Volume("*mm^3*")"), 
					picfile = paste0("Abs_TV_", gv, ".png"), legendMargin = FALSE, zeroStart = TRUE, yStart = 0)
}

################################################
### 'Mean ± SEM' plot of %Change Body Weight ###
################################################
# 合并体重改变数据的 平均值(mean) 和 标准误差(sem) 为一个数据框(dataframe)
BWdePer4Plot <- MergeMeanSEM(meanData = BWmeanPerData, semData = BWsemPerData)
# 为每个组别添加颜色和形状，并设置为factor
BWdePer4PlotList <- addColShape(dataFrame = BWdePer4Plot)
BWdePer4newPlot <- BWdePer4PlotList$data
colorBWperLevels <- BWdePer4PlotList$color
shapeBWperLevels <- BWdePer4PlotList$shape
# 将GroupList中每个组别的 Treatment 添加到 BWdePer4newPlot$group
BWdePer4newPlot = AddTreat2DataFrameV2(dataF = BWdePer4newPlot, treatInfo = GroupList)				
# plot
PlotMeanSEM(data = BWdePer4newPlot, maintitle = "%Change Body Weight Mean ± SEM", colS = colorBWperLevels, shapeS = shapeBWperLevels,
					xlab = "Treatment Days", ylab = "% Change BW(%)", picName = "BWdePer4Plot.png")

###################################
###### ends of plots ##############
###################################

#####################################
#####################################
######## write analysis result ######
#####################################
#####################################

wb <- openxlsx::createWorkbook()
openxlsx::modifyBaseFont(wb, fontSize = 10)
####################
## 'Report'工作簿 ##
####################
sheet1 <- openxlsx::addWorksheet(wb, sheetName="Report", gridLines = FALSE)
subStyle = openxlsx::createStyle(fontSize = 20, textDecoration = "bold", halign="center", valign="center")
# header
openxlsx::mergeCells(wb, sheet1, cols=1:12, rows=11)
openxlsx::writeData(wb, sheet1, x="Coherentbio efficacy report (PDX model)", startCol=1, startRow = 11)
openxlsx::addStyle(wb, sheet1, subStyle, rows = 11, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet1, rows=11, heights=25)
# study ID
for(i in 13:17){
	openxlsx::mergeCells(wb, sheet1, cols=3:4, rows=i)
	openxlsx::addStyle(wb, sheet1, createStyle(halign="right"), rows = i, cols = 3, gridExpand = FALSE, stack = TRUE)
	openxlsx::mergeCells(wb, sheet1, cols=5:11, rows=i)
	openxlsx::addStyle(wb, sheet1, createStyle(textDecoration = "bold"), rows = i, cols = 5, gridExpand = FALSE, stack = TRUE)
	if(i == 13){
		openxlsx::writeData(wb, sheet1, "Study:", startCol = 3, startRow = 13, rowNames = FALSE, colNames = FALSE)
		openxlsx::writeData(wb, sheet1, projectID, startCol = 5, startRow = 13, rowNames = FALSE, colNames = FALSE)
	}else if(i == 14){
		openxlsx::writeData(wb, sheet1, "Date Range:", startCol = 3, startRow = 14, rowNames = FALSE, colNames = FALSE)
		openxlsx::writeData(wb, sheet1, "Dosing and Recovery Phases", startCol = 5, startRow = 14, rowNames = FALSE, colNames = FALSE)
	}else if(i == 15){
		openxlsx::writeData(wb, sheet1, "Generated by:", startCol = 3, startRow = 15, rowNames = FALSE, colNames = FALSE)
		openxlsx::writeData(wb, sheet1, StudyDirector, startCol = 5, startRow = 15, rowNames = FALSE, colNames = FALSE)
	}else if(i == 16){
		openxlsx::writeData(wb, sheet1, "Generated on:", startCol = 3, startRow = 16, rowNames = FALSE, colNames = FALSE)
		openxlsx::writeData(wb, sheet1, format(Sys.time(), "%d/%m/%Y %H:%M"), startCol = 5, startRow = 16, rowNames = FALSE, colNames = FALSE)
	}else if(i == 17){
		openxlsx::writeData(wb, sheet1, "Version:", startCol = 3, startRow = 17, rowNames = FALSE, colNames = FALSE)
		openxlsx::writeData(wb, sheet1, "v1.0.0.0", startCol = 5, startRow = 17, rowNames = FALSE, colNames = FALSE)
	}
}
####################
## 'Index'工作簿 ###
####################
sheet2 <- openxlsx::addWorksheet(wb, sheetName="Index", gridLines = FALSE)
subStyle1 <- openxlsx::createStyle(fontSize = 12, textDecoration = "bold", halign="left", valign="center")
hs1 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","white","white"))

## subject infos ##
openxlsx::mergeCells(wb, sheet2, cols=1:4, rows=1)
openxlsx::writeData(wb, sheet2, x="General Information", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet2, createStyle(fontSize = 18, textDecoration = "bold", halign="left", valign="center"), rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet2, rows=1, heights=20)
openxlsx::setColWidths(wb, sheet2, cols = 4:6, widths = 30)
for(i in 2:4){
	openxlsx::mergeCells(wb, sheet2, cols=1:3, rows=i)
	openxlsx::mergeCells(wb, sheet2, cols=4:15, rows=i)
}
subData1 <- data.frame(matrix(rep("", 45),nrow=3,ncol=15))
subData1$X1 = c("Study Number:", "Title:", "Objective:")
subData1$X4 = c(projectID, projectTitle, Object)
openxlsx::writeData(wb, sheet2, subData1, startCol = 1, startRow = 2, rowNames = FALSE, colNames = FALSE, borders = "surrounding", borderColour = "black")
textLeftAlign = openxlsx::createStyle(halign="right", valign="center")
openxlsx::addStyle(wb, sheet2, textLeftAlign, rows = 2:4, cols = 1, gridExpand = FALSE, stack = TRUE)
## personel ##
openxlsx::mergeCells(wb, sheet2, cols=1:3, rows=7)
openxlsx::writeData(wb, sheet2, x="Personel", startCol=1, startRow = 7)
openxlsx::conditionalFormatting(wb, sheet2, cols=1, rows=7, rule = "!=0", style = subStyle1)
openxlsx::setRowHeights(wb, sheet2, rows=7, heights = 15)

personelCount = 2 + length(LabTechList)
personelEndLine = 8 + personelCount

for(ii in 8:personelEndLine){
	openxlsx::mergeCells(wb, sheet2, cols=1:3, rows=ii)
	openxlsx::mergeCells(wb, sheet2, cols=5:6, rows=ii)
}
# 将Personel数据写入dataframe
subData2 <- data.frame(matrix(rep("", personelCount*6),nrow=personelCount,ncol=6))
colnames(subData2)[1] = "Role"
colnames(subData2)[4] = "Name"
colnames(subData2)[5] = "Contact"
subData2$Role = c("Study Director", "Technician Group Leader", rep("Lab Technician",length(LabTechList)))
subData2$Name = c(StudyDirector[1], TechGroupLeader[1], sapply(seq(1:length(LabTechList)), function(x){LabTechList[[x]][1]}))
subData2$Contact = c(StudyDirector[2], TechGroupLeader[2], sapply(seq(1:length(LabTechList)), function(x){LabTechList[[x]][2]}))
# 写入工作簿
openxlsx::writeData(wb, sheet2, subData2, startCol = 1, startRow = 8, rowNames = FALSE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::addStyle(wb, sheet2, createStyle(border=c("left"), borderColour="black"), rows = 8, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet2, createStyle(border=c("right"), borderColour="black"), rows = 8, cols = 6, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet2, createStyle(halign = "CENTER",valign = "CENTER"), rows = 9:personelEndLine, cols = 2:6, gridExpand = TRUE, stack = TRUE)
## Treatment ##
TreatStartLine = personelEndLine + 3
openxlsx::mergeCells(wb, sheet2, cols=1:3, rows=TreatStartLine)
openxlsx::writeData(wb, sheet2, x="Treatments", startCol=1, startRow = TreatStartLine)
openxlsx::conditionalFormatting(wb, sheet2, cols=1, rows=TreatStartLine, rule = "!=0", style = subStyle1)
openxlsx::setRowHeights(wb, sheet2, rows=TreatStartLine, heights = 15)

TreatCount = sum(sapply(seq(1:length(GroupList)), function(x){length(GroupList[[x]])}))
for(iii in (TreatStartLine + 1):(TreatStartLine + 1 + TreatCount)){
	openxlsx::mergeCells(wb, sheet2, cols=1:3, rows=iii)
	openxlsx::mergeCells(wb, sheet2, cols=4:6, rows=iii)
}
# 将GroupList数据写入dataframe
GroupListDf = ldply(GroupList, data.frame)
# 将dataframe修改成适合写入有格式的单元格
subData3 <- data.frame(matrix(rep("", TreatCount*6),nrow=TreatCount,ncol=6))
colnames(subData3)[1] = "Group"
colnames(subData3)[4] = "Treatment"
subData3$Group = GroupListDf[,1]
subData3$Treatment = GroupListDf[,2]
writeData(wb, sheet2, subData3, startCol = 1, startRow = TreatStartLine + 1, rowNames = FALSE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
addStyle(wb, sheet2, createStyle(border=c("left"), borderColour="black"), rows = TreatStartLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet2, createStyle(border=c("right"), borderColour="black"), rows = TreatStartLine + 1, cols = 6, gridExpand = TRUE, stack = TRUE)
		
############################
## 'Dosing record'工作簿 ###
############################
sheet3 <- openxlsx::addWorksheet(wb, sheetName="Dosing record", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet3, cols = 6:(length(dayName)+5), widths = 14)
openxlsx::setColWidths(wb, sheet3, cols = 1:5, widths = 12)
# 设置工作簿主标题Style
subStyle3 <- openxlsx::createStyle(fontSize = 16, textDecoration = "bold", halign="center", valign="center")
# 设置小标题Style
subStyle4 <- openxlsx::createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3", halign="center", valign="center")
# 设置列表表头Style
hs2 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","black","black"))
# 写入工作簿主标题
openxlsx::mergeCells(wb, sheet3, cols=1:6, rows=1)
openxlsx::writeData(wb, sheet3, x="Dosing record", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet3, subStyle3, rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet3, rows=1, heights=20)
# 写入研究ID
openxlsx::mergeCells(wb, sheet3, cols=1:2, rows=2)
openxlsx::writeData(wb, sheet3, x="Study Number:", startCol=1, startRow = 2)
openxlsx::mergeCells(wb, sheet3, cols=3:6, rows=2)
openxlsx::writeData(wb, sheet3, x=projectID, startCol=3, startRow = 2)
# 写入表例
openxlsx::writeData(wb, sheet3, x="Legend", startCol=1, startRow = 4)
openxlsx::addStyle(wb, sheet3, createStyle(textDecoration = "bold"), rows = 4, cols = 1, gridExpand = FALSE, stack = TRUE)
LegendStartLine = 5
for(Li in 1:(length(GroupTreatData$Group) + 1)){
	openxlsx::mergeCells(wb, sheet3, cols = 2:15, rows = LegendStartLine + Li - 1)
}
newGroupTreatData = cbind(GroupTreatData, matrix(nrow=length(GroupTreatData$Group), ncol = 13, byrow = TRUE))
openxlsx::writeData(wb, sheet3, newGroupTreatData, startCol = 1, startRow = LegendStartLine, rowNames = FALSE, colNames = TRUE, 
			borders = "columns", borderColour = "black", headerStyle = hs2)
openxlsx::addStyle(wb, sheet3, createStyle(fontSize = 8), rows = (LegendStartLine + 1):(LegendStartLine + length(GroupTreatData$Group)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
# 开始写入Dosing Record
DoseStartLine = LegendStartLine + length(GroupTreatData$Group) + 2
openxlsx::mergeCells(wb, sheet3, cols = 6:(length(dayName)+5), rows = DoseStartLine)
openxlsx::writeData(wb, sheet3, x = "Dates/Study Days", startCol = 6, startRow = DoseStartLine)
openxlsx::addStyle(wb, sheet3, subStyle4, rows = DoseStartLine, cols = 6:(length(dayName)+5), gridExpand = FALSE, stack = TRUE)
# 写入有效记录日期
for(j in 1:length(dayName)){
	#prjnt(j)
	openxlsx::writeData(wb, sheet3, x=paste("Date:", dateName[j]), startCol=(j-1)+6, startRow = DoseStartLine + 1)
	openxlsx::addStyle(wb, sheet3, createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3"), rows = DoseStartLine + 1, cols = (j-1)+6, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet3, x=paste("Days:", dayName[j]), startCol=(j-1)+6, startRow = DoseStartLine + 2)
	openxlsx::addStyle(wb, sheet3, createStyle(border = c("top", "bottom", "left", "right"), fgFill = "#D3D3D3"), rows = DoseStartLine + 2, cols = (j-1)+6, gridExpand = TRUE, stack = TRUE)
}
# 写入组别、用药和样本ID的列名
openxlsx::writeData(wb, sheet3, x="Group", startCol=1, startRow = DoseStartLine + 2)
openxlsx::addStyle(wb, sheet3, subStyle4, rows = DoseStartLine + 2, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::mergeCells(wb, sheet3, cols=2:4, rows=DoseStartLine + 2)
openxlsx::writeData(wb, sheet3, x="Treatment", startCol=2, startRow = DoseStartLine + 2)
openxlsx::addStyle(wb, sheet3, subStyle4, rows = DoseStartLine + 2, cols = 2:4, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet3, x="SampleID", startCol=5, startRow = DoseStartLine + 2)
openxlsx::addStyle(wb, sheet3, subStyle4, rows = DoseStartLine + 2, cols = 5, gridExpand = TRUE, stack = TRUE)
# 根据 GroupListDf 和 SampleList ，写入组别名称、用药情况
DoseGroupstartRowID = DoseStartLine + 3
DoseGroupTmpID = DoseGroupstartRowID
for(k in 1:nrow(GroupListDf)){
	tmpGroupName = GroupListDf[k,1]
	openxlsx::mergeCells(wb, sheet3, cols=1, rows=DoseGroupTmpID:(DoseGroupTmpID + length(SampleList[[tmpGroupName]]) - 1))
	openxlsx::addStyle(wb, sheet3, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = DoseGroupTmpID:(DoseGroupTmpID + length(SampleList[[tmpGroupName]]) - 1), cols = 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet3, x=tmpGroupName, startCol=1, startRow = DoseGroupTmpID)
	openxlsx::mergeCells(wb, sheet3, cols=2:4, rows=DoseGroupTmpID:(DoseGroupTmpID + length(SampleList[[tmpGroupName]]) - 1))
	openxlsx::addStyle(wb, sheet3, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center", wrapText=TRUE), rows = DoseGroupTmpID:(DoseGroupTmpID + length(SampleList[[tmpGroupName]]) - 1), cols = 2:4, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet3, x=as.character(GroupListDf[k,2]), startCol=2, startRow = DoseGroupTmpID)
	#addStyle(wb, sheet3, createStyle(border = c("bottom", "left", "top", "right"),halign="center", valign="center"), rows = DoseGroupTmpID:(DoseGroupTmpID + length(SampleList[[tmpGroupName]]) - 1), cols = 5, gridExpand = TRUE, stack = TRUE)
	DoseGroupTmpID = DoseGroupTmpID + length(SampleList[[tmpGroupName]])
}
# 根据SampleRecord，写入样本ID和有效记录
openxlsx::writeData(wb, sheet3, SampleRecord, startCol=5, startRow = DoseGroupstartRowID, colNames = FALSE, rowNames = FALSE, 
			borders = "none")
# 设置记录字符居中显示
openxlsx::addStyle(wb, sheet3, createStyle(halign="center", valign="center"), rows = DoseGroupstartRowID:(DoseGroupTmpID - 1), cols = 5:(length(dayName)+5), gridExpand = TRUE, stack = TRUE)			
# 为DosingRecord区域添加边框
for(ji in 0:length(dayName)){
	openxlsx::addStyle(wb, sheet3, createStyle(border="right"), rows = DoseGroupstartRowID:(DoseGroupTmpID - 1), cols = 5+ji, gridExpand = TRUE, stack = TRUE)
}
DoseGroupTmpID2 = DoseGroupstartRowID
for(jk in 1:nrow(GroupListDf)){
	tmpGroupName = GroupListDf[jk,1]
	openxlsx::addStyle(wb, sheet3, createStyle(border="bottom", borderStyle = "dashed"), rows = DoseGroupTmpID2:(DoseGroupTmpID2 + length(SampleList[[tmpGroupName]]) - 2), cols = 6:(length(dayName)+5), gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet3, createStyle(border="bottom"), rows = DoseGroupTmpID2 + length(SampleList[[tmpGroupName]]) - 1, cols = 5:(length(dayName)+5), gridExpand = TRUE, stack = TRUE)
	DoseGroupTmpID2 = DoseGroupTmpID2 + length(SampleList[[tmpGroupName]])
}

noteData = matrix(c("Note", "√", "dosed", "", "Note", "N/A", "No dosing per protocol", "", "Note", "Holiday", "Holiday", ""), nrow=3, byrow=T)
openxlsx::mergeCells(wb, sheet3, cols=3:4, rows=DoseGroupTmpID2 + 4)
openxlsx::mergeCells(wb, sheet3, cols=3:4, rows=DoseGroupTmpID2 + 5)
openxlsx::mergeCells(wb, sheet3, cols=3:4, rows=DoseGroupTmpID2 + 6)
openxlsx::writeData(wb, sheet3, noteData, startCol=1, startRow = DoseGroupTmpID2 + 4, colNames = FALSE, rowNames = FALSE, 
			borders = "columns", borderColour = "black", headerStyle = hs1)
openxlsx::mergeCells(wb, sheet3, cols=1, rows=(DoseGroupTmpID2 + 4):(DoseGroupTmpID2 + 6))
openxlsx::addStyle(wb, sheet3, createStyle(border="bottom"), rows = (DoseGroupTmpID2 + 4):(DoseGroupTmpID2 + 6), cols = 1:4, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet3, createStyle(halign="center",valign="center",border="bottom"), rows = (DoseGroupTmpID2 + 4):(DoseGroupTmpID2 + 6), cols = 1:2, gridExpand = TRUE, stack = TRUE)

############################
### 'Sample List'工作簿 ####
############################
sheet4 <- openxlsx::addWorksheet(wb, sheetName="Sample list", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet4, cols = 3:(length(solidList)+5), widths = 14)
openxlsx::setColWidths(wb, sheet4, cols = 1:2, widths = 12)

# 写入工作簿主标题
openxlsx::mergeCells(wb, sheet4, cols=1:6, rows=1)
openxlsx::writeData(wb, sheet4, x="Sample list", startCol=1, startRow = 1)
openxlsx::addStyle(wb, sheet4, subStyle3, rows = 1, cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::setRowHeights(wb, sheet4, rows=1, heights=20)
# 写入研究ID
openxlsx::mergeCells(wb, sheet4, cols=1:2, rows=2)
openxlsx::writeData(wb, sheet4, x="Study Number:", startCol=1, startRow = 2)
openxlsx::mergeCells(wb, sheet4, cols=3:6, rows=2)
openxlsx::writeData(wb, sheet4, x=projectID, startCol=3, startRow = 2)
# 写入表例
openxlsx::writeData(wb, sheet4, x="Legend", startCol=1, startRow = 4)
openxlsx::addStyle(wb, sheet4, createStyle(textDecoration = "bold"), rows = 4, cols = 1, gridExpand = FALSE, stack = TRUE)
LegendStartLine1 = 5
for(sli in 1:(length(GroupTreatData$Group) + 1)){
	openxlsx::mergeCells(wb, sheet4, cols = 2:15, rows = LegendStartLine1 + sli - 1)
}
# newGroupTreatData沿用“Dosing record”工作簿中的图例
openxlsx::writeData(wb, sheet4, newGroupTreatData, startCol = 1, startRow = LegendStartLine1, rowNames = FALSE, colNames = TRUE, 
			borders = "columns", borderColour = "black", headerStyle = hs2)
openxlsx::addStyle(wb, sheet4, createStyle(fontSize = 8), rows = (LegendStartLine1 + 1):(LegendStartLine1 + length(GroupTreatData$Group)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
# 开始写入Sample list
SampleLiStartLine = LegendStartLine1 + length(GroupTreatData$Group) + 2
headerList = c("Group", "SampleID", solidList)
for(sL in 1:length(headerList)){
	openxlsx::writeData(wb, sheet4, x=headerList[sL], startCol=sL, startRow = SampleLiStartLine + 1)
}
openxlsx::addStyle(wb, sheet4, subStyle4, rows = SampleLiStartLine + 1, cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)

SampleDissect <- NULL
for(ss in 1:length(SDlist)){
	if(is.null(SampleDissect)){
		SampleDissect = SDlist[[ss]][['data']]
	}else{
		SampleDissect = rbind(SampleDissect, SDlist[[ss]][['data']])
	}
}
openxlsx::writeData(wb, sheet4, SampleDissect, startCol = 3, startRow = SampleLiStartLine + 2, rowNames = FALSE, colNames = FALSE, 
			borders = "rows", borderColour = "black", borderStyle="dashed")
openxlsx::addStyle(wb, sheet4, createStyle(halign="center", valign="center"), rows = (SampleLiStartLine + 2):(SampleLiStartLine + 2+nrow(SampleDissect) - 1), cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)			
for(hi in 1:length(headerList)){
	openxlsx::addStyle(wb, sheet4, createStyle(border="right"), rows = (SampleLiStartLine + 2):(SampleLiStartLine + 2+nrow(SampleDissect) - 1), cols = hi, gridExpand = TRUE, stack = TRUE)
}
preGroupLine = 0
afterGroupLine = 0
for(hk in 1:length(SDlist)){
	tmpSampleIndex = which(rownames(SampleDissect) %in% SDlist[[hk]][['samples']])
	afterGroupLine = preGroupLine + length(tmpSampleIndex)
	openxlsx::mergeCells(wb, sheet4, cols=1, rows=(SampleLiStartLine + 2 + preGroupLine):(SampleLiStartLine + 2 + afterGroupLine - 1))
	openxlsx::addStyle(wb, sheet4, createStyle(border="bottom"), rows = SampleLiStartLine + 2 + afterGroupLine - 1, cols = 1:length(headerList), gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheet4, x=names(SDlist)[hk], startCol=1, startRow = SampleLiStartLine+2+preGroupLine)
	for(hkk in 1:length(tmpSampleIndex)){
		openxlsx::writeData(wb, sheet4, x=SDlist[[hk]][['samples']][hkk], startCol=2, startRow = SampleLiStartLine + 2 +tmpSampleIndex[hkk] - 1)
		openxlsx::addStyle(wb, sheet4, createStyle(border="bottom"), rows = SampleLiStartLine + 2+tmpSampleIndex[hkk] - 1, cols = 2, gridExpand = TRUE, stack = TRUE)
	}
	preGroupLine = afterGroupLine
}


############################
### 'Tumor volume'工作簿 ###
############################
sheet5 <- openxlsx::addWorksheet(wb, sheetName="Tumor volume", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet5, cols = 1:(length(dayName)+5+5), widths = 12)

# 写入Subject ID
openxlsx::mergeCells(wb, sheet5, cols=1:2, rows=1)
openxlsx::mergeCells(wb, sheet5, cols=3:6, rows=1)
openxlsx::mergeCells(wb, sheet5, cols=3:6, rows=2)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = 2, cols = 3, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="Study Number:", startCol=1, startRow = 1)
openxlsx::writeData(wb, sheet5, x=projectID, startCol=3, startRow = 1)
openxlsx::writeData(wb, sheet5, x="Tumor Volume (mm3): ", startCol=3, startRow = 2)
# 写入表例
openxlsx::writeData(wb, sheet5, x="Legend", startCol=1, startRow = 4)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = 4, cols = 1, gridExpand = FALSE, stack = TRUE)
LegendStartLine2 = 5
for(Li2 in 1:(length(GroupTreatData$Group) + 1)){
	openxlsx::mergeCells(wb, sheet5, cols = 2:15, rows = LegendStartLine2 + Li2 - 1)
}
# newGroupTreatData沿用“Dosing record”工作簿中的图例
openxlsx::writeData(wb, sheet5, newGroupTreatData, startCol = 1, startRow = LegendStartLine2, rowNames = FALSE, colNames = TRUE, 
			borders = "columns", borderColour = "black", headerStyle = hs2)
openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 8), rows = (LegendStartLine2 + 1):(LegendStartLine2 + length(GroupTreatData$Group)), cols = 1:2, gridExpand = TRUE, stack = TRUE)

# 添加点线图：肿瘤体积的'Mean ± SEM'
TVdeStartLine = LegendStartLine2 + length(GroupTreatData$Group) + 2
openxlsx::insertImage(wb, sheet5, "TVde4Plot.png", width = 12, height = 3, startRow = TVdeStartLine, startCol = 1, dpi=400)
TVdeEndLine = TVdeStartLine + 18
# 开始写入Mean ± SEM data of Tumor volume
# TVdeDataStartLine = TVdeEndLine
openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=TVdeEndLine)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = TVdeEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="Group Mean +/- Standard Error of the Mean, Tumor Volume (mm3): ", startCol=1, startRow = TVdeEndLine)
TVdeDataStartLine = TVdeEndLine + 2
TVdeDataEndLine = TVdeDataStartLine + nrow(TVmeanData)
# 写入数据框mean of Tumor Volume
openxlsx::writeData(wb, sheet5, TVmeanData, startCol = 1, startRow = TVdeDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet5, x="Group", startCol = 1, startRow = TVdeDataStartLine)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left"), borderColour="black"), rows = TVdeDataStartLine:TVdeDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = TVdeDataStartLine, cols = ncol(TVmeanData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = TVdeDataStartLine:TVdeDataEndLine, cols = 1:(ncol(TVmeanData) + 1), gridExpand = TRUE, stack = TRUE)
# 准备格式
openxlsx::mergeCells(wb, sheet5, cols=1:4, rows=TVdeDataEndLine + 2)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = TVdeDataEndLine + 2, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="StdErr Mean", startCol=1, startRow = TVdeDataEndLine + 2)
TVsemDataStartLine = TVdeDataEndLine + 4
TVsemDataEndLine = TVsemDataStartLine + nrow(TVsemData)
# 写入数据框SEM of Tumor Volume
openxlsx::writeData(wb, sheet5, TVsemData, startCol = 1, startRow = TVsemDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet5, x="Group", startCol = 1, startRow = TVsemDataStartLine)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left"), borderColour="black"), rows = TVsemDataStartLine:TVsemDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = TVsemDataStartLine, cols = ncol(TVsemData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = TVsemDataStartLine:TVsemDataEndLine, cols = 1:(ncol(TVsemData) + 1), gridExpand = TRUE, stack = TRUE)

# 添加点线图：%T/C Tumor Volume
TCperStartLine = TVsemDataEndLine + 2
openxlsx::insertImage(wb, sheet5, "TVtc4Plot.png", width = 12, height = 3, startRow = TCperStartLine, startCol = 1, dpi=400)
TCperEndLine = TCperStartLine + 18
# 准备格式
openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=TCperEndLine + 1)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = TCperEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="% T/C, Tumor Volume", startCol=1, startRow = TCperEndLine + 1)
TCperDataStartLine = TCperEndLine + 2
TCperData <- extractDataInfo(FromList = TVdata, nameCols = "%T/C", Formula = NULL)
TCperDataEndLine = TCperDataStartLine + nrow(TCperData) + 2
# 写入数据框%T/C Tumor Volume
# 设置表头
openxlsx::mergeCells(wb, sheet5, cols=2:(ncol(TCperData) + 1), rows=TCperDataStartLine)
openxlsx::writeData(wb, sheet5, x = "Dates/Study Days", startCol = 2, startRow = TCperDataStartLine)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = TCperDataStartLine, cols = 2:(ncol(TCperData) + 1), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet5, cols=1, rows=(TCperDataStartLine + 1):(TCperDataStartLine + 2))
openxlsx::writeData(wb, sheet5, x = "Group", startCol = 1, startRow = TCperDataStartLine + 1)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = (TCperDataStartLine + 1):(TCperDataStartLine + 2), cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::writeData(wb, sheet5, matrix(recordDate, nrow=1), startCol = 2, startRow = TCperDataStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet5, matrix(as.numeric(recordDay), nrow=1), startCol = 2, startRow = TCperDataStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (TCperDataStartLine + 1):(TCperDataStartLine + 2), cols = 2:(ncol(TCperData) + 1), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = (TCperDataStartLine + 1):(TCperDataStartLine + 2), cols = ncol(TCperData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (TCperDataStartLine + 1):(TCperDataStartLine + 2), cols = 2:(ncol(TCperData) + 1), gridExpand = TRUE, stack = TRUE)
# 写入数据框%T/C
openxlsx::writeData(wb, sheet5, TCperData, startCol = 1, startRow = TCperDataStartLine + 3, rowNames = TRUE, colNames = FALSE, 
			borders = "surrounding", borderColour = "black")
openxlsx::addStyle(wb, sheet5, createStyle(border=c("RIGHT"), borderColour="black"), rows = (TCperDataStartLine + 3):TCperDataEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = (TCperDataStartLine + 3):TCperDataEndLine, cols = 1:(ncol(TCperData) + 1), gridExpand = TRUE, stack = TRUE)
# 添加计算公式及说明
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCperDataEndLine + 1)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCperDataEndLine + 2)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCperDataEndLine + 3)
openxlsx::writeData(wb, sheet5, x="%T/C = mean(T)/mean(C) * 100%", startCol = 1, startRow = TCperDataEndLine + 1)
openxlsx::writeData(wb, sheet5, x="T - current group value", startCol = 1, startRow = TCperDataEndLine + 2)
openxlsx::writeData(wb, sheet5, x="C - control group value", startCol = 1, startRow = TCperDataEndLine + 3)
openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 7, textDecoration = "italic"), rows = (TCperDataEndLine + 1):(TCperDataEndLine + 3), cols = 1, gridExpand = TRUE, stack = TRUE)

# 添加点线图：%Inhibition Tumor Volume
InhPerStartLine = TCperDataEndLine + 5
openxlsx::insertImage(wb, sheet5, "TVInh4Plot.png", width = 12, height = 3, startRow = InhPerStartLine, startCol = 1, dpi=400)
InhPerEndLine = InhPerStartLine + 18
# 准备格式
openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=InhPerEndLine + 1)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = InhPerEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="Mean % Inhibition, Tumor Volume", startCol=1, startRow = InhPerEndLine + 1)
InhPerDataStartLine = InhPerEndLine + 2
InhPerData <- extractDataInfo(FromList = TVdata, nameCols = "%Inhibition", Formula = NULL)
InhPerDataEndLine = InhPerDataStartLine + nrow(InhPerData) + 2
# 写入数据框%Inhibition Tumor Volume
# 设置表头
openxlsx::mergeCells(wb, sheet5, cols=2:(ncol(InhPerData) + 1), rows=InhPerDataStartLine)
openxlsx::writeData(wb, sheet5, x = "Dates/Study Days", startCol = 2, startRow = InhPerDataStartLine)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = InhPerDataStartLine, cols = 2:(ncol(InhPerData) + 1), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet5, cols=1, rows=(InhPerDataStartLine + 1):(InhPerDataStartLine + 2))
openxlsx::writeData(wb, sheet5, x = "Group", startCol = 1, startRow = InhPerDataStartLine + 1)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = (InhPerDataStartLine + 1):(InhPerDataStartLine + 2), cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::writeData(wb, sheet5, matrix(recordDate, nrow=1), startCol = 2, startRow = InhPerDataStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet5, matrix(as.numeric(recordDay), nrow=1), startCol = 2, startRow = InhPerDataStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (InhPerDataStartLine + 1):(InhPerDataStartLine + 2), cols = 2:(ncol(InhPerData) + 1), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = (InhPerDataStartLine + 1):(InhPerDataStartLine + 2), cols = ncol(InhPerData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (InhPerDataStartLine + 1):(InhPerDataStartLine + 2), cols = 2:(ncol(InhPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 写入数据框%Inhibition
openxlsx::writeData(wb, sheet5, InhPerData, startCol = 1, startRow = InhPerDataStartLine + 3, rowNames = TRUE, colNames = FALSE, 
			borders = "surrounding", borderColour = "black")
openxlsx::addStyle(wb, sheet5, createStyle(border=c("RIGHT"), borderColour="black"), rows = (InhPerDataStartLine + 3):InhPerDataEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = (InhPerDataStartLine + 3):InhPerDataEndLine, cols = 1:(ncol(InhPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 添加计算公式及说明
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhPerDataEndLine + 1)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhPerDataEndLine + 2)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhPerDataEndLine + 3)
openxlsx::writeData(wb, sheet5, x="Mean % Inhibition = (mean(C)-mean(T))/mean(C) * 100%", startCol = 1, startRow = InhPerDataEndLine + 1)
openxlsx::writeData(wb, sheet5, x="T - current group value", startCol = 1, startRow = InhPerDataEndLine + 2)
openxlsx::writeData(wb, sheet5, x="C - control group value", startCol = 1, startRow = InhPerDataEndLine + 3)
openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 7, textDecoration = "italic"), rows = (InhPerDataEndLine + 1):(InhPerDataEndLine + 3), cols = 1, gridExpand = TRUE, stack = TRUE)

# 添加点线图：%ΔT/ΔC Tumor Volume
TCdeltaPerStartLine = InhPerDataEndLine + 5
openxlsx::insertImage(wb, sheet5, "TVtcDeltaPer4Plot.png", width = 12, height = 3, startRow = TCdeltaPerStartLine, startCol = 1, dpi=400)
TCdeltaPerEndLine = TCdeltaPerStartLine + 18
# 准备格式
openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=TCdeltaPerEndLine + 1)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = TCdeltaPerEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="% ΔT/ΔC, Tumor Volume", startCol=1, startRow = TCdeltaPerEndLine + 1)
TCdeltaPerDataStartLine = TCdeltaPerEndLine + 2
TCdeltaPerData <- extractDataInfo(FromList = TVdata, nameCols = "%ΔT/ΔC", Formula = NULL)
TCdeltaPerDataEndLine = TCdeltaPerDataStartLine + nrow(TCdeltaPerData) + 2
# 写入数据框%ΔT/ΔC Tumor Volume
# 设置表头
openxlsx::mergeCells(wb, sheet5, cols=2:(ncol(TCdeltaPerData) + 1), rows=TCdeltaPerDataStartLine)
openxlsx::writeData(wb, sheet5, x = "Dates/Study Days", startCol = 2, startRow = TCdeltaPerDataStartLine)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = TCdeltaPerDataStartLine, cols = 2:(ncol(TCdeltaPerData) + 1), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet5, cols=1, rows=(TCdeltaPerDataStartLine + 1):(TCdeltaPerDataStartLine + 2))
openxlsx::writeData(wb, sheet5, x = "Group", startCol = 1, startRow = TCdeltaPerDataStartLine + 1)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = (TCdeltaPerDataStartLine + 1):(TCdeltaPerDataStartLine + 2), cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::writeData(wb, sheet5, matrix(recordDate, nrow=1), startCol = 2, startRow = TCdeltaPerDataStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet5, matrix(as.numeric(recordDay), nrow=1), startCol = 2, startRow = TCdeltaPerDataStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (TCdeltaPerDataStartLine + 1):(TCdeltaPerDataStartLine + 2), cols = 2:(ncol(TCdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = (TCdeltaPerDataStartLine + 1):(TCdeltaPerDataStartLine + 2), cols = ncol(TCdeltaPerData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (TCdeltaPerDataStartLine + 1):(TCdeltaPerDataStartLine + 2), cols = 2:(ncol(TCdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 写入数据框%ΔT/ΔC
openxlsx::writeData(wb, sheet5, TCdeltaPerData, startCol = 1, startRow = TCdeltaPerDataStartLine + 3, rowNames = TRUE, colNames = FALSE, 
			borders = "surrounding", borderColour = "black")
openxlsx::addStyle(wb, sheet5, createStyle(border=c("RIGHT"), borderColour="black"), rows = (TCdeltaPerDataStartLine + 3):TCdeltaPerDataEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = (TCdeltaPerDataStartLine + 3):TCdeltaPerDataEndLine, cols = 1:(ncol(TCdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 添加计算公式及说明
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCdeltaPerDataEndLine + 1)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCdeltaPerDataEndLine + 2)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCdeltaPerDataEndLine + 3)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCdeltaPerDataEndLine + 4)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=TCdeltaPerDataEndLine + 5)
openxlsx::writeData(wb, sheet5, x="% ΔT/ΔC = (mean(T)-mean(T0)) / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = TCdeltaPerDataEndLine + 1)
openxlsx::writeData(wb, sheet5, x="T - current group value", startCol = 1, startRow = TCdeltaPerDataEndLine + 2)
openxlsx::writeData(wb, sheet5, x="T0 - current group initial value", startCol = 1, startRow = TCdeltaPerDataEndLine + 3)
openxlsx::writeData(wb, sheet5, x="C - control group value", startCol = 1, startRow = TCdeltaPerDataEndLine + 4)
openxlsx::writeData(wb, sheet5, x="C0 - control group initial value", startCol = 1, startRow = TCdeltaPerDataEndLine + 5)
openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 7, textDecoration = "italic"), rows = (TCdeltaPerDataEndLine + 1):(TCdeltaPerDataEndLine + 5), cols = 1, gridExpand = TRUE, stack = TRUE)

# 添加点线图：%ΔInhibition Tumor Volume
InhdeltaPerStartLine = TCdeltaPerDataEndLine + 7
openxlsx::insertImage(wb, sheet5, "TVInhDeltaPer4Plot.png", width = 12, height = 3, startRow = InhdeltaPerStartLine, startCol = 1, dpi=400)
InhdeltaPerEndLine = InhdeltaPerStartLine + 18
# 准备格式
openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=InhdeltaPerEndLine + 1)
openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = InhdeltaPerEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet5, x="Mean % ΔInhibition, Tumor Volume", startCol=1, startRow = InhdeltaPerEndLine + 1)
InhdeltaPerDataStartLine = InhdeltaPerEndLine + 2
InhdeltaPerData <- extractDataInfo(FromList = TVdata, nameCols = "%Δinhibition", Formula = NULL)
InhdeltaPerDataEndLine = InhdeltaPerDataStartLine + nrow(InhdeltaPerData) + 2
# 写入数据框%ΔInhibition Tumor Volume
# 设置表头
openxlsx::mergeCells(wb, sheet5, cols=2:(ncol(InhdeltaPerData) + 1), rows=InhdeltaPerDataStartLine)
openxlsx::writeData(wb, sheet5, x = "Dates/Study Days", startCol = 2, startRow = InhdeltaPerDataStartLine)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = InhdeltaPerDataStartLine, cols = 2:(ncol(InhdeltaPerData) + 1), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet5, cols=1, rows=(InhdeltaPerDataStartLine + 1):(InhdeltaPerDataStartLine + 2))
openxlsx::writeData(wb, sheet5, x = "Group", startCol = 1, startRow = InhdeltaPerDataStartLine + 1)
openxlsx::addStyle(wb, sheet5, subStyle4, rows = (InhdeltaPerDataStartLine + 1):(InhdeltaPerDataStartLine + 2), cols = 1, gridExpand = FALSE, stack = TRUE)
openxlsx::writeData(wb, sheet5, matrix(recordDate, nrow=1), startCol = 2, startRow = InhdeltaPerDataStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet5, matrix(as.numeric(recordDay), nrow=1), startCol = 2, startRow = InhdeltaPerDataStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (InhdeltaPerDataStartLine + 1):(InhdeltaPerDataStartLine + 2), cols = 2:(ncol(InhdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = (InhdeltaPerDataStartLine + 1):(InhdeltaPerDataStartLine + 2), cols = ncol(InhdeltaPerData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (InhdeltaPerDataStartLine + 1):(InhdeltaPerDataStartLine + 2), cols = 2:(ncol(InhdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 写入数据框%ΔT/ΔC
openxlsx::writeData(wb, sheet5, InhdeltaPerData, startCol = 1, startRow = InhdeltaPerDataStartLine + 3, rowNames = TRUE, colNames = FALSE, 
			borders = "surrounding", borderColour = "black")
openxlsx::addStyle(wb, sheet5, createStyle(border=c("RIGHT"), borderColour="black"), rows = (InhdeltaPerDataStartLine + 3):InhdeltaPerDataEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = (InhdeltaPerDataStartLine + 3):InhdeltaPerDataEndLine, cols = 1:(ncol(InhdeltaPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 添加计算公式及说明
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhdeltaPerDataEndLine + 1)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhdeltaPerDataEndLine + 2)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhdeltaPerDataEndLine + 3)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhdeltaPerDataEndLine + 4)
openxlsx::mergeCells(wb, sheet5, cols=1:6, rows=InhdeltaPerDataEndLine + 5)
openxlsx::writeData(wb, sheet5, x="Mean % ΔInhibition = ((mean(C)-mean(C0)) - (mean(T)-mean(T0))) / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = InhdeltaPerDataEndLine + 1)
openxlsx::writeData(wb, sheet5, x="T - current group value", startCol = 1, startRow = InhdeltaPerDataEndLine + 2)
openxlsx::writeData(wb, sheet5, x="T0 - current group initial value", startCol = 1, startRow = InhdeltaPerDataEndLine + 3)
openxlsx::writeData(wb, sheet5, x="C - control group value", startCol = 1, startRow = InhdeltaPerDataEndLine + 4)
openxlsx::writeData(wb, sheet5, x="C0 - control group initial value", startCol = 1, startRow = InhdeltaPerDataEndLine + 5)
openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 7, textDecoration = "italic"), rows = (InhdeltaPerDataEndLine + 1):(InhdeltaPerDataEndLine + 5), cols = 1, gridExpand = TRUE, stack = TRUE)

### 分组计算结果
grStartLine = InhdeltaPerDataEndLine + 8
hs1 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","white","white"))

for(gm in 1:length(TVdataFinal)){
	openxlsx::mergeCells(wb, sheet5, cols = 1:10, rows = grStartLine + 1)
	openxlsx::writeData(wb, sheet5, x=GroupTreatData$Caption[which(GroupTreatData$Group == names(TVdataFinal)[gm])], startCol = 1, startRow = grStartLine + 1)
	openxlsx::addStyle(wb, sheet5, createStyle(textDecoration = "bold"), rows = grStartLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
	grEndLine = grStartLine + 3 + nrow(TVdataFinal[[gm]])
	openxlsx::writeData(wb, sheet5, TVdataFinal[[gm]], startCol = 1, startRow = grStartLine + 3, rowNames = TRUE, colNames = TRUE,
				borders = "columns", borderColour = "black", headerStyle = hs1)
	openxlsx::addStyle(wb, sheet5, createStyle(border=c("left"), borderColour="black"), rows = grStartLine + 3, cols = 1:2, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet5, createStyle(border=c("right"), borderColour="black"), rows = grStartLine + 3, cols = ncol(TVdataFinal[[gm]]) + 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet5, createStyle(border=c("bottom"), borderColour="black"), rows = grStartLine + 3 + nrow(TVdata[[gm]]$data), cols = 1:(ncol(TVdataFinal[[gm]]) + 1), gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet5, createStyle(halign = "CENTER",valign = "CENTER"), rows = (grStartLine + 3):grEndLine, cols = 1:(ncol(TVdataFinal[[gm]]) + 1), gridExpand = TRUE, stack = TRUE)
	# 添加计算公式及说明
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 1)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 2)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 3)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 4)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 5)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 6)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 7)
	openxlsx::mergeCells(wb, sheet5, cols=1:10, rows=grEndLine + 8)
	openxlsx::writeData(wb, sheet5, x="%T/C = mean(T)/mean(C) * 100%", startCol = 1, startRow = grEndLine + 1)
	openxlsx::writeData(wb, sheet5, x="%Inhibition = (mean(C)-mean(T)) / mean(C) * 100%", startCol = 1, startRow = grEndLine + 2)
	openxlsx::writeData(wb, sheet5, x="%ΔT/ΔC = (mean(T)-mean(T0)) / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = grEndLine + 3)
	openxlsx::writeData(wb, sheet5, x="%ΔInhibition = ((mean(C)-mean(C0)) - (mean(T)-mean(T0)))  / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = grEndLine + 4)
	openxlsx::writeData(wb, sheet5, x="T - current group value", startCol = 1, startRow = grEndLine + 5)
	openxlsx::writeData(wb, sheet5, x="T0 - current group initial value", startCol = 1, startRow = grEndLine + 6)
	openxlsx::writeData(wb, sheet5, x="C - control group value", startCol = 1, startRow = grEndLine + 7)
	openxlsx::writeData(wb, sheet5, x="C0 - control group initial value", startCol = 1, startRow = grEndLine + 8)
	openxlsx::addStyle(wb, sheet5, createStyle(fontSize = 7, textDecoration = "italic"), rows = (grEndLine + 1):(grEndLine + 8), cols = 1, gridExpand = TRUE, stack = TRUE)

	grStartLine = grStartLine + 3 + nrow(TVdataFinal[[gm]]) + 10
}

############################
### 'Body weight'工作簿 ####
############################
sheet6 <- openxlsx::addWorksheet(wb, sheetName="Body weight", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet6, cols = 1:(length(dayName)+5+5), widths = 10)

# 写入Subject ID
openxlsx::mergeCells(wb, sheet6, cols=1:2, rows=1)
openxlsx::mergeCells(wb, sheet6, cols=3:6, rows=1)
openxlsx::mergeCells(wb, sheet6, cols=3:6, rows=2)
openxlsx::addStyle(wb, sheet6, createStyle(textDecoration = "bold"), rows = 2, cols = 3, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet6, x="Study Number:", startCol=1, startRow = 1)
openxlsx::writeData(wb, sheet6, x=projectID, startCol=3, startRow = 1)
openxlsx::writeData(wb, sheet6, x="Body Weight (g)", startCol=3, startRow = 2)
# 写入表例
openxlsx::writeData(wb, sheet6, x="Legend", startCol=1, startRow = 4)
openxlsx::addStyle(wb, sheet6, createStyle(textDecoration = "bold"), rows = 4, cols = 1, gridExpand = FALSE, stack = TRUE)
LegendStartLine3 = 5
for(Li3 in 1:(length(GroupTreatData$Group) + 1)){
	openxlsx::mergeCells(wb, sheet6, cols = 2:15, rows = LegendStartLine3 + Li3 - 1)
}
# newGroupTreatData沿用“Dosing record”工作簿中的图例
openxlsx::writeData(wb, sheet6, newGroupTreatData, startCol = 1, startRow = LegendStartLine3, rowNames = FALSE, colNames = TRUE, 
			borders = "columns", borderColour = "black", headerStyle = hs2)
openxlsx::addStyle(wb, sheet6, createStyle(fontSize = 8), rows = (LegendStartLine3 + 1):(LegendStartLine3 + length(GroupTreatData$Group)), cols = 1:2, gridExpand = TRUE, stack = TRUE)

# 添加点线图：模型体重的'Mean ± SEM'
BWdeStartLine = LegendStartLine3 + length(GroupTreatData$Group) + 2
openxlsx::insertImage(wb, sheet6, "BWde4Plot.png", width = 12, height = 3, startRow = BWdeStartLine, startCol = 1, dpi=400)
BWdeEndLine = BWdeStartLine + 18
# 开始写入Mean ± SEM data of Body weight
# BWdeDataStartLine = BWdeEndLine
openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=BWdeEndLine)
openxlsx::addStyle(wb, sheet6, createStyle(textDecoration = "bold"), rows = BWdeEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet6, x="Group Mean +/- Standard Error of the Mean, Body Weight (g)", startCol=1, startRow = BWdeEndLine)
BWdeDataStartLine = BWdeEndLine + 2
BWdeDataEndLine = BWdeDataStartLine + nrow(BWmeanData)
# 写入数据框mean of Body weight
openxlsx::writeData(wb, sheet6, BWmeanData, startCol = 1, startRow = BWdeDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet6, x="Group", startCol = 1, startRow = BWdeDataStartLine)
openxlsx::addStyle(wb, sheet6, createStyle(border=c("left"), borderColour="black"), rows = BWdeDataStartLine:BWdeDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet6, createStyle(border=c("right"), borderColour="black"), rows = BWdeDataStartLine, cols = ncol(BWmeanData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet6, createStyle(halign = "CENTER",valign = "CENTER"), rows = BWdeDataStartLine:BWdeDataEndLine, cols = 1:(ncol(BWmeanData) + 1), gridExpand = TRUE, stack = TRUE)
# 准备格式
openxlsx::mergeCells(wb, sheet6, cols=1:4, rows=BWdeDataEndLine + 2)
openxlsx::addStyle(wb, sheet6, createStyle(textDecoration = "bold"), rows = BWdeDataEndLine + 2, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet6, x="StdErr Mean", startCol=1, startRow = BWdeDataEndLine + 2)
BWsemDataStartLine = BWdeDataEndLine + 4
BWsemDataEndLine = BWsemDataStartLine + nrow(BWsemData)
# 写入数据框SEM of Body weight
openxlsx::writeData(wb, sheet6, BWsemData, startCol = 1, startRow = BWsemDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet6, x="Group", startCol = 1, startRow = BWsemDataStartLine)
openxlsx::addStyle(wb, sheet6, createStyle(border=c("left"), borderColour="black"), rows = BWsemDataStartLine:BWsemDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet6, createStyle(border=c("right"), borderColour="black"), rows = BWsemDataStartLine, cols = ncol(BWsemData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet6, createStyle(halign = "CENTER",valign = "CENTER"), rows = BWsemDataStartLine:BWsemDataEndLine, cols = 1:(ncol(BWsemData) + 1), gridExpand = TRUE, stack = TRUE)

### 分组计算结果
grBWStartLine = BWsemDataEndLine + 4
hs1 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","white","white"))

for(gw in 1:length(BWdataFinal)){
	openxlsx::mergeCells(wb, sheet6, cols = 1:10, rows = grBWStartLine + 1)
	openxlsx::writeData(wb, sheet6, x=GroupTreatData$Caption[which(GroupTreatData$Group == names(BWdataFinal)[gw])], startCol = 1, startRow = grBWStartLine + 1)
	openxlsx::addStyle(wb, sheet6, createStyle(textDecoration = "bold"), rows = grBWStartLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
	grBWEndLine = grBWStartLine + 3 + nrow(BWdataFinal[[gw]])
	openxlsx::writeData(wb, sheet6, BWdataFinal[[gw]], startCol = 1, startRow = grBWStartLine + 3, rowNames = TRUE, colNames = TRUE,
				borders = "columns", borderColour = "black", headerStyle = hs1)
	openxlsx::addStyle(wb, sheet6, createStyle(border=c("left"), borderColour="black"), rows = grBWStartLine + 3, cols = 1:2, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet6, createStyle(border=c("right"), borderColour="black"), rows = grBWStartLine + 3, cols = ncol(BWdataFinal[[gw]]) + 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet6, createStyle(border=c("bottom"), borderColour="black"), rows = grBWStartLine + 3 + nrow(BWdata[[gw]]$data), cols = 1:(ncol(BWdataFinal[[gw]]) + 1), gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet6, createStyle(halign = "CENTER",valign = "CENTER"), rows = (grBWStartLine + 3):grBWEndLine, cols = 1:(ncol(BWdataFinal[[gw]]) + 1), gridExpand = TRUE, stack = TRUE)
	# 添加计算公式及说明
	openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=grBWEndLine + 1)
	openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=grBWEndLine + 2)
	openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=grBWEndLine + 3)
	openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=grBWEndLine + 4)
	openxlsx::mergeCells(wb, sheet6, cols=1:10, rows=grBWEndLine + 5)
	openxlsx::writeData(wb, sheet6, x="%T/C = mean(T)/mean(C) * 100%", startCol = 1, startRow = grBWEndLine + 1)
	openxlsx::writeData(wb, sheet6, x="%Inhibition = (mean(C)-mean(T)) / mean(C) * 100%", startCol = 1, startRow = grBWEndLine + 2)
	openxlsx::writeData(wb, sheet6, x="%ΔT/ΔC = (mean(T)-mean(T0)) / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = grBWEndLine + 3)
	openxlsx::writeData(wb, sheet6, x="%ΔInhibition = ((mean(C)-mean(C0)) - (mean(T)-mean(T0)))  / (mean(C)-mean(C0)) * 100%", startCol = 1, startRow = grBWEndLine + 4)
	openxlsx::writeData(wb, sheet6, x="T - current group value", startCol = 1, startRow = grBWEndLine + 5)
	openxlsx::addStyle(wb, sheet6, createStyle(fontSize = 7, textDecoration = "italic"), rows = (grBWEndLine + 1):(grBWEndLine + 5), cols = 1, gridExpand = TRUE, stack = TRUE)

	grBWStartLine = grBWStartLine + 3 + nrow(BWdataFinal[[gw]]) + 7
}

############################
### 'BW Change %'工作簿 ####
############################
sheet7 <- openxlsx::addWorksheet(wb, sheetName="BW Change %", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet7, cols = 1:(length(dayName)+5+5), widths = 10)

# 写入Subject ID
openxlsx::mergeCells(wb, sheet7, cols=1:2, rows=1)
openxlsx::mergeCells(wb, sheet7, cols=3:6, rows=1)
openxlsx::mergeCells(wb, sheet7, cols=3:6, rows=2)
openxlsx::addStyle(wb, sheet7, createStyle(textDecoration = "bold"), rows = 2, cols = 3, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet7, x="Study Number:", startCol=1, startRow = 1)
openxlsx::writeData(wb, sheet7, x=projectID, startCol=3, startRow = 1)
openxlsx::writeData(wb, sheet7, x="Body Weight Change", startCol=3, startRow = 2)
# 写入表例
openxlsx::writeData(wb, sheet7, x="Legend", startCol=1, startRow = 4)
openxlsx::addStyle(wb, sheet7, createStyle(textDecoration = "bold"), rows = 4, cols = 1, gridExpand = FALSE, stack = TRUE)
LegendStartLine4 = 5
for(Li4 in 1:(length(GroupTreatData$Group) + 1)){
	openxlsx::mergeCells(wb, sheet7, cols = 2:15, rows = LegendStartLine4 + Li4 - 1)
}
# newGroupTreatData沿用“Dosing record”工作簿中的图例
openxlsx::writeData(wb, sheet7, newGroupTreatData, startCol = 1, startRow = LegendStartLine4, rowNames = FALSE, colNames = TRUE, 
			borders = "columns", borderColour = "black", headerStyle = hs2)
openxlsx::addStyle(wb, sheet7, createStyle(fontSize = 8), rows = (LegendStartLine4 + 1):(LegendStartLine4 + length(GroupTreatData$Group)), cols = 1:2, gridExpand = TRUE, stack = TRUE)

# 添加点线图：模型改变体重的'Mean ± SEM'
BWChangedeStartLine = LegendStartLine4 + length(GroupTreatData$Group) + 2
openxlsx::insertImage(wb, sheet7, "BWdePer4Plot.png", width = 12, height = 3, startRow = BWChangedeStartLine, startCol = 1, dpi=400)
BWChangedeEndLine = BWChangedeStartLine + 18
# 开始写入Mean ± SEM data of Body weight change
# BWChangedeDataStartLine = BWChangedeEndLine
openxlsx::mergeCells(wb, sheet7, cols=1:10, rows=BWChangedeEndLine)
openxlsx::addStyle(wb, sheet7, createStyle(textDecoration = "bold"), rows = BWChangedeEndLine, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet7, x="% Group Change Mean +/- Standard Error of the Mean, Body Weight", startCol=1, startRow = BWChangedeEndLine)
BWChangedeDataStartLine = BWChangedeEndLine + 2
BWChangedeDataEndLine = BWChangedeDataStartLine + nrow(BWmeanPerData)
# 写入数据框mean of Body weight change
openxlsx::writeData(wb, sheet7, BWmeanPerData, startCol = 1, startRow = BWChangedeDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet7, x="Group", startCol = 1, startRow = BWChangedeDataStartLine)
openxlsx::addStyle(wb, sheet7, createStyle(border=c("left"), borderColour="black"), rows = BWChangedeDataStartLine:BWChangedeDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet7, createStyle(border=c("right"), borderColour="black"), rows = BWChangedeDataStartLine, cols = ncol(BWmeanPerData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet7, createStyle(halign = "CENTER",valign = "CENTER"), rows = BWChangedeDataStartLine:BWChangedeDataEndLine, cols = 1:(ncol(BWmeanPerData) + 1), gridExpand = TRUE, stack = TRUE)
# 准备格式
openxlsx::mergeCells(wb, sheet7, cols=1:4, rows=BWChangedeDataEndLine + 2)
openxlsx::addStyle(wb, sheet7, createStyle(textDecoration = "bold"), rows = BWChangedeDataEndLine + 2, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet7, x="StdErr Mean", startCol=1, startRow = BWChangedeDataEndLine + 2)
BWChangesemDataStartLine = BWChangedeDataEndLine + 4
BWChangesemDataEndLine = BWChangesemDataStartLine + nrow(BWsemPerData)
# 写入数据框SEM of Body weight
openxlsx::writeData(wb, sheet7, BWsemPerData, startCol = 1, startRow = BWChangesemDataStartLine, rowNames = TRUE, colNames = TRUE, 
			borders = "surrounding", borderColour = "black", headerStyle = hs1)
openxlsx::writeData(wb, sheet7, x="Group", startCol = 1, startRow = BWChangesemDataStartLine)
openxlsx::addStyle(wb, sheet7, createStyle(border=c("left"), borderColour="black"), rows = BWChangesemDataStartLine:BWChangesemDataEndLine, cols = 1:2, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet7, createStyle(border=c("right"), borderColour="black"), rows = BWChangesemDataStartLine, cols = ncol(BWsemPerData) + 1, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet7, createStyle(halign = "CENTER",valign = "CENTER"), rows = BWChangesemDataStartLine:BWChangesemDataEndLine, cols = 1:(ncol(BWsemPerData) + 1), gridExpand = TRUE, stack = TRUE)

### 分组计算结果
grBWChangeStartLine = BWChangesemDataEndLine + 4
hs1 <- openxlsx::createStyle(fgFill = "#D3D3D3", halign = "CENTER", border=c("top","left","right") , borderColour=c("black","white","white"))

for(gh in 1:length(BWChangedataFinal)){
	openxlsx::mergeCells(wb, sheet7, cols = 1:10, rows = grBWChangeStartLine + 1)
	openxlsx::writeData(wb, sheet7, x=GroupTreatData$Caption[which(GroupTreatData$Group == names(BWChangedataFinal)[gh])], startCol = 1, startRow = grBWChangeStartLine + 1)
	openxlsx::addStyle(wb, sheet7, createStyle(textDecoration = "bold"), rows = grBWChangeStartLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
	grBWChangeEndLine = grBWChangeStartLine + 3 + nrow(BWChangedataFinal[[gh]])
	openxlsx::writeData(wb, sheet7, BWChangedataFinal[[gh]], startCol = 1, startRow = grBWChangeStartLine + 3, rowNames = TRUE, colNames = TRUE,
				borders = "columns", borderColour = "black", headerStyle = hs1)
	openxlsx::addStyle(wb, sheet7, createStyle(border=c("left"), borderColour="black"), rows = grBWChangeStartLine + 3, cols = 1:2, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet7, createStyle(border=c("right"), borderColour="black"), rows = grBWChangeStartLine + 3, cols = ncol(BWChangedataFinal[[gh]]) + 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet7, createStyle(border=c("bottom"), borderColour="black"), rows = grBWChangeStartLine + 3 + nrow(BWChangedata[[gh]]$data), cols = 1:(ncol(BWChangedataFinal[[gh]]) + 1), gridExpand = TRUE, stack = TRUE)
	openxlsx::addStyle(wb, sheet7, createStyle(halign = "CENTER",valign = "CENTER"), rows = (grBWChangeStartLine + 3):grBWChangeEndLine, cols = 1:(ncol(BWChangedataFinal[[gh]]) + 1), gridExpand = TRUE, stack = TRUE)

	grBWChangeStartLine = grBWChangeStartLine + 3 + nrow(BWChangedataFinal[[gh]]) + 2
}

#################################
#### 'Data summary BW'工作簿 ####
#################################
sheet8 <- openxlsx::addWorksheet(wb, sheetName="Data summary BW", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet8, cols = 1:(length(dayName)+5+5), widths = 12)

# 写入工作簿主标题
openxlsx::mergeCells(wb, sheet8, cols=1:10, rows=1)
openxlsx::mergeCells(wb, sheet8, cols=1:10, rows=2)
openxlsx::mergeCells(wb, sheet8, cols=1:10, rows=3)
openxlsx::addStyle(wb, sheet8, createStyle(fontSize = 12, textDecoration = "bold"), rows = 1:3, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet8, x=projectID, startCol=1, startRow = 1)
openxlsx::writeData(wb, sheet8, x="Task: Body Weight (g)", startCol=1, startRow = 2)
openxlsx::writeData(wb, sheet8, x="Individual Measurements", startCol=1, startRow = 3)
openxlsx::addStyle(wb, sheet8, subStyle3, rows = 1:3, cols = 1, gridExpand = FALSE, stack = TRUE)
# 写入小标题
openxlsx::mergeCells(wb, sheet8, cols=1:4, rows=6)
openxlsx::addStyle(wb, sheet8, createStyle(textDecoration = "bold"), rows = 6, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet8, x="Absolute Body Weight (g)", startCol=1, startRow = 6)

# 合并所有样本的Body weight数据
BWsumData <- extractDataInfo(FromList = BWdata, nameCols = "data", Formula = NULL, rowNameChange = FALSE)
SampleGroupData <- NULL
for(gr in 1:length(SampleList)){
	grTname = names(SampleList)[gr]
	tmpDF = data.frame(Group = rep(grTname, length(SampleList[[gr]])), samples = SampleList[[gr]])
	if(is.null(SampleGroupData)){
		SampleGroupData = tmpDF
	}else{
		SampleGroupData = rbind(SampleGroupData, tmpDF)
	}
}
FateData <- NULL
for(sm in 1:nrow(BWsumData)){
	tmpDayCount = length(na.omit(BWsumData[sm,]))
	tmpCode = "A"
	if(tmpDayCount < ncol(BWsumData)){
		tmpCode = "TS"
	}
	if(is.null(FateData)){
		FateData = c(tmpCode, tmpDayCount)
	}else{
		FateData = rbind(FateData, c(tmpCode, tmpDayCount))
	}
}
rownames(FateData) = seq(1, nrow(FateData))
FateData = as.data.frame(FateData, stringsAsFactors = FALSE)
FateData$V2 = as.numeric(FateData$V2)

BWsumStartLine = 7
# 设置表头
openxlsx::mergeCells(wb, sheet8, cols=5:(ncol(BWsumData) + 4), rows=BWsumStartLine)
openxlsx::writeData(wb, sheet8, x = "Dates/Study Days", startCol = 5, startRow = BWsumStartLine)
openxlsx::addStyle(wb, sheet8, subStyle4, rows = BWsumStartLine, cols = 5:(ncol(BWsumData) + 4), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet8, cols=1, rows=(BWsumStartLine + 1):(BWsumStartLine + 2))
openxlsx::writeData(wb, sheet8, x = "Group", startCol = 1, startRow = BWsumStartLine + 1)
openxlsx::mergeCells(wb, sheet8, cols=2, rows=(BWsumStartLine + 1):(BWsumStartLine + 2))
openxlsx::writeData(wb, sheet8, x = "Sample ID", startCol = 2, startRow = BWsumStartLine + 1)
openxlsx::mergeCells(wb, sheet8, cols=3:4, rows = BWsumStartLine + 1)
openxlsx::writeData(wb, sheet8, x = "Fate", startCol = 3, startRow = BWsumStartLine + 1)
openxlsx::writeData(wb, sheet8, x = "Code", startCol = 3, startRow = BWsumStartLine + 2)
openxlsx::writeData(wb, sheet8, x = "Day", startCol = 4, startRow = BWsumStartLine + 2)
openxlsx::writeData(wb, sheet8, matrix(recordDate, nrow=1), startCol = 5, startRow = BWsumStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet8, matrix(as.numeric(recordDay), nrow=1), startCol = 5, startRow = BWsumStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet8, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (BWsumStartLine + 1):(BWsumStartLine + 2), cols = 5:(ncol(BWsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet8, createStyle(border=c("right"), borderColour="black"), rows = (BWsumStartLine + 1):(BWsumStartLine + 2), cols = ncol(BWsumData) + 4, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet8, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (BWsumStartLine + 1):(BWsumStartLine + 2), cols = 1:(ncol(BWsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet8, createStyle(border=c("top","right","bottom","left"), borderColour=rep("black",4)), rows = (BWsumStartLine + 1):(BWsumStartLine + 2), cols = 1:4, gridExpand = TRUE, stack = TRUE)
# 写入数据
if((length(unique(rownames(BWsumData) == as.character(unlist(SampleList)))) == 1) && (unique(rownames(BWsumData) == as.character(unlist(SampleList))))){
	if(unique(rownames(BWsumData) == as.character(unlist(SampleList)))){
		openxlsx::writeData(wb, sheet8, SampleGroupData, startCol = 1, startRow = BWsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
		openxlsx::writeData(wb, sheet8, BWsumData, startCol = 5, startRow = BWsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
		openxlsx::writeData(wb, sheet8, FateData, startCol = 3, startRow = BWsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
	}
}else{
	stop("sampleinfos in SampleList are different from BWsumData in sheet \'Data summary BW\'")
}
BWTmpStartLine = BWsumStartLine + 3
for(grt in 1:length(BWdata)){
	openxlsx::addStyle(wb, sheet8, createStyle(border=c("bottom"), borderColour="black"), 
				rows = BWTmpStartLine + nrow(BWdata[[grt]]$data) - 1, cols = 1:(ncol(BWsumData) + 4), gridExpand = TRUE, stack = TRUE)
	BWTmpStartLine = BWTmpStartLine + nrow(BWdata[[grt]]$data)
}
openxlsx::addStyle(wb, sheet8, createStyle(halign = "center", valign = "center"), rows = (BWsumStartLine + 3):(BWsumStartLine + 2 + nrow(BWsumData)), 
				cols = 1:(ncol(BWsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet8, createStyle(border = "right", borderColour = "black"), rows = (BWsumStartLine + 3):(BWsumStartLine + 2 + nrow(BWsumData)), 
				cols = c(1,3), gridExpand = TRUE, stack = TRUE)
BWsumEndLine = BWsumStartLine + 2 + nrow(BWsumData)
# 添加注释
openxlsx::mergeCells(wb, sheet8, cols=1:6, rows=BWsumEndLine + 1)
openxlsx::writeData(wb, sheet8, x="TS is Terminal Sacrifice ", startCol = 1, startRow = BWsumEndLine + 1)
openxlsx::addStyle(wb, sheet8, createStyle(fontSize = 7, textDecoration = "italic"), rows = BWsumEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
# 插入图片
dayBWStartLine = BWsumEndLine + 4
for(gbp in 1:length(BWdata)){
	openxlsx::insertImage(wb, sheet8, paste0("Abs_BW_", gbp, ".png"), width = 7, height = 2.6, startRow = dayBWStartLine, startCol = 1, dpi=400)
	dayBWStartLine = dayBWStartLine + 18
}

#################################
#### 'Data summary TV'工作簿 ####
#################################
sheet9 <- openxlsx::addWorksheet(wb, sheetName="Data summary TV", gridLines = FALSE)
# 设置列宽
openxlsx::setColWidths(wb, sheet9, cols = 1:(length(dayName)+5+5), widths = 12)

# 写入工作簿主标题
openxlsx::mergeCells(wb, sheet9, cols=1:10, rows=1)
openxlsx::mergeCells(wb, sheet9, cols=1:10, rows=2)
openxlsx::mergeCells(wb, sheet9, cols=1:10, rows=3)
openxlsx::addStyle(wb, sheet9, createStyle(fontSize = 12, textDecoration = "bold"), rows = 1:3, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet9, x=projectID, startCol=1, startRow = 1)
openxlsx::writeData(wb, sheet9, x="Task: Tumor Volume (mm3): ", startCol=1, startRow = 2)
openxlsx::writeData(wb, sheet9, x="Individual Measurements", startCol=1, startRow = 3)
openxlsx::addStyle(wb, sheet9, subStyle3, rows = 1:3, cols = 1, gridExpand = FALSE, stack = TRUE)
# 写入小标题
openxlsx::mergeCells(wb, sheet9, cols=1:4, rows=6)
openxlsx::addStyle(wb, sheet9, createStyle(textDecoration = "bold"), rows = 6, cols = 1, gridExpand = TRUE, stack = TRUE)
openxlsx::writeData(wb, sheet9, x="Absolute Tumor Volume (mm3): ", startCol=1, startRow = 6)

# 合并所有样本的Tumor Volume数据
TVsumData <- extractDataInfo(FromList = TVdata, nameCols = "data", Formula = NULL, rowNameChange = FALSE)
tvFateData <- NULL
for(sv in 1:nrow(TVsumData)){
	tmpDayCount = length(na.omit(TVsumData[sv,]))
	tmpCode = "A"
	if(tmpDayCount < ncol(TVsumData)){
		tmpCode = "TS"
	}
	if(is.null(tvFateData)){
		tvFateData = c(tmpCode, tmpDayCount)
	}else{
		tvFateData = rbind(tvFateData, c(tmpCode, tmpDayCount))
	}
}
rownames(tvFateData) = seq(1, nrow(tvFateData))
tvFateData = as.data.frame(tvFateData, stringsAsFactors = FALSE)
tvFateData$V2 = as.numeric(tvFateData$V2)

TVsumStartLine = 7
# 设置表头
openxlsx::mergeCells(wb, sheet9, cols=5:(ncol(TVsumData) + 4), rows=TVsumStartLine)
openxlsx::writeData(wb, sheet9, x = "Dates/Study Days", startCol = 5, startRow = TVsumStartLine)
openxlsx::addStyle(wb, sheet9, subStyle4, rows = TVsumStartLine, cols = 5:(ncol(TVsumData) + 4), gridExpand = FALSE, stack = TRUE)
openxlsx::mergeCells(wb, sheet9, cols=1, rows=(TVsumStartLine + 1):(TVsumStartLine + 2))
openxlsx::writeData(wb, sheet9, x = "Group", startCol = 1, startRow = TVsumStartLine + 1)
openxlsx::mergeCells(wb, sheet9, cols=2, rows=(TVsumStartLine + 1):(TVsumStartLine + 2))
openxlsx::writeData(wb, sheet9, x = "Sample ID", startCol = 2, startRow = TVsumStartLine + 1)
openxlsx::mergeCells(wb, sheet9, cols=3:4, rows = TVsumStartLine + 1)
openxlsx::writeData(wb, sheet9, x = "Fate", startCol = 3, startRow = TVsumStartLine + 1)
openxlsx::writeData(wb, sheet9, x = "Code", startCol = 3, startRow = TVsumStartLine + 2)
openxlsx::writeData(wb, sheet9, x = "Day", startCol = 4, startRow = TVsumStartLine + 2)
openxlsx::writeData(wb, sheet9, matrix(recordDate, nrow=1), startCol = 5, startRow = TVsumStartLine + 1, rowNames = FALSE, colNames = FALSE)
openxlsx::writeData(wb, sheet9, matrix(as.numeric(recordDay), nrow=1), startCol = 5, startRow = TVsumStartLine + 2, rowNames = FALSE, colNames = FALSE)
openxlsx::addStyle(wb, sheet9, createStyle(border=c("left", "bottom"), borderColour=c("white","black")), rows = (TVsumStartLine + 1):(TVsumStartLine + 2), cols = 5:(ncol(TVsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet9, createStyle(border=c("right"), borderColour="black"), rows = (TVsumStartLine + 1):(TVsumStartLine + 2), cols = ncol(TVsumData) + 4, gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet9, createStyle(fgFill = "#D3D3D3", halign = "CENTER",valign = "CENTER"), rows = (TVsumStartLine + 1):(TVsumStartLine + 2), cols = 1:(ncol(TVsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet9, createStyle(border=c("top","right","bottom","left"), borderColour=rep("black",4)), rows = (TVsumStartLine + 1):(TVsumStartLine + 2), cols = 1:4, gridExpand = TRUE, stack = TRUE)
# 写入数据
if((length(unique(rownames(TVsumData) == as.character(unlist(SampleList)))) == 1) && (unique(rownames(TVsumData) == as.character(unlist(SampleList))))){
	if(unique(rownames(TVsumData) == as.character(unlist(SampleList)))){
		openxlsx::writeData(wb, sheet9, SampleGroupData, startCol = 1, startRow = TVsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
		openxlsx::writeData(wb, sheet9, TVsumData, startCol = 5, startRow = TVsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
		openxlsx::writeData(wb, sheet9, tvFateData, startCol = 3, startRow = TVsumStartLine + 3, colNames = FALSE, rowNames = FALSE,
					borders = "surrounding", borderColour = "black")
	}
}else{
	stop("sampleinfos in SampleList are different from TVsumData in sheet \'Data summary TV\'")
}
TVTmpStartLine = TVsumStartLine + 3
for(gvt in 1:length(TVdata)){
	openxlsx::addStyle(wb, sheet9, createStyle(border=c("bottom"), borderColour="black"), 
				rows = TVTmpStartLine + nrow(TVdata[[gvt]]$data) - 1, cols = 1:(ncol(TVsumData) + 4), gridExpand = TRUE, stack = TRUE)
	TVTmpStartLine = TVTmpStartLine + nrow(TVdata[[gvt]]$data)
}
openxlsx::addStyle(wb, sheet9, createStyle(halign = "center", valign = "center"), rows = (TVsumStartLine + 3):(TVsumStartLine + 2 + nrow(TVsumData)), 
				cols = 1:(ncol(TVsumData) + 4), gridExpand = TRUE, stack = TRUE)
openxlsx::addStyle(wb, sheet9, createStyle(border = "right", borderColour = "black"), rows = (TVsumStartLine + 3):(TVsumStartLine + 2 + nrow(TVsumData)), 
				cols = c(1,3), gridExpand = TRUE, stack = TRUE)
TVsumEndLine = TVsumStartLine + 2 + nrow(TVsumData)
# 添加注释
openxlsx::mergeCells(wb, sheet9, cols=1:6, rows=TVsumEndLine + 1)
openxlsx::writeData(wb, sheet9, x="TS is Terminal Sacrifice ", startCol = 1, startRow = TVsumEndLine + 1)
openxlsx::addStyle(wb, sheet9, createStyle(fontSize = 7, textDecoration = "italic"), rows = TVsumEndLine + 1, cols = 1, gridExpand = TRUE, stack = TRUE)
# 插入图片
dayTVStartLine = TVsumEndLine + 4
for(gvp in 1:length(TVdata)){
	openxlsx::insertImage(wb, sheet9, paste0("Abs_TV_", gvp, ".png"), width = 7, height = 2.6, startRow = dayTVStartLine, startCol = 1, dpi=400)
	dayTVStartLine = dayTVStartLine + 18
}

#################################
####### 'Day *****'工作簿 #######
#################################
# merge Length of Tumor Volume
TVlenMergeData = extractDataInfo(FromList = TVLendata, nameCols = "Length", Formula = NULL, rowNameChange = FALSE)
# merge Width of Tumor Volume
TVwidMergeData = extractDataInfo(FromList = TVWiddata, nameCols = "Width", Formula = NULL, rowNameChange = FALSE)

for(d in 1:length(recordDay)){
	# d = 1
	tmpDay = as.character(recordDay)[d]
	sheetTmp <- openxlsx::addWorksheet(wb, sheetName=paste("Day", tmpDay), gridLines = FALSE)
	# 设置列宽
	openxlsx::setColWidths(wb, sheetTmp, cols = 1:26, widths = 10)
	# 设置行高
	openxlsx::setRowHeights(wb, sheetTmp, rows = 7, heights = 15)

	# 写入工作簿主标题
	openxlsx::mergeCells(wb, sheetTmp, cols=1:10, rows=1)
	openxlsx::mergeCells(wb, sheetTmp, cols=1:10, rows=3)
	openxlsx::mergeCells(wb, sheetTmp, cols=1:10, rows=4)
	openxlsx::addStyle(wb, sheetTmp, createStyle(fontSize = 20, textDecoration = "bold", halign="center", valign="center"), rows = 1, cols = 1, gridExpand = TRUE, stack = TRUE)
	openxlsx::writeData(wb, sheetTmp, x=projectID, startCol=1, startRow = 1)
	openxlsx::writeData(wb, sheetTmp, x="All Tasks", startCol=1, startRow = 3)
	openxlsx::writeData(wb, sheetTmp, x=paste0("Day ", as.character(recordDay)[d], ", ", as.character(recordDate)[d]), startCol=1, startRow = 4)
	openxlsx::addStyle(wb, sheetTmp, createStyle(fontSize = 14, textDecoration = "bold", halign="center", valign="center"), rows = 3:4, cols = 1, gridExpand = FALSE, stack = TRUE)
	
	tmpDayStartLine = 6
	# 设置表头
	# Group, Sample ID, Cage
	openxlsx::mergeCells(wb, sheetTmp, cols=1, rows=tmpDayStartLine:(tmpDayStartLine + 1))
	openxlsx::writeData(wb, sheetTmp, x = "Group", startCol = 1, startRow = tmpDayStartLine)
	openxlsx::mergeCells(wb, sheetTmp, cols=2, rows=tmpDayStartLine:(tmpDayStartLine + 1))
	openxlsx::writeData(wb, sheetTmp, x = "Sample ID", startCol = 2, startRow = tmpDayStartLine)
	openxlsx::mergeCells(wb, sheetTmp, cols=3, rows=tmpDayStartLine:(tmpDayStartLine + 1))
	openxlsx::writeData(wb, sheetTmp, x = "Cage", startCol = 3, startRow = tmpDayStartLine)
	# Body Weight
	openxlsx::mergeCells(wb, sheetTmp, cols=4:5, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Body Weight (g)", startCol = 4, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Weight", startCol = 4, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 5, startRow = tmpDayStartLine + 1)
	# Tumor Volume
	openxlsx::mergeCells(wb, sheetTmp, cols=6:9, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Tumor Volume (mm3)", startCol = 6, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Length", startCol = 6, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Width", startCol = 7, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Volume", startCol = 8, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 9, startRow = tmpDayStartLine + 1)
	# Article 1
	openxlsx::mergeCells(wb, sheetTmp, cols=10:11, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Article 1", startCol = 10, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Name", startCol = 10, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 11, startRow = tmpDayStartLine + 1)
	# Article 2
	openxlsx::mergeCells(wb, sheetTmp, cols=12:13, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Article 2", startCol = 12, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Name", startCol = 12, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 13, startRow = tmpDayStartLine + 1)
	# Article 3
	openxlsx::mergeCells(wb, sheetTmp, cols=14:15, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Article 3", startCol = 14, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Name", startCol = 14, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 15, startRow = tmpDayStartLine + 1)
	# Clinical Observation
	openxlsx::mergeCells(wb, sheetTmp, cols=16:17, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Clinical Observation", startCol = 16, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Observations", startCol = 16, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 17, startRow = tmpDayStartLine + 1)
	# Mortality Observation
	openxlsx::mergeCells(wb, sheetTmp, cols=18:19, rows=tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Mortality Observation", startCol = 18, startRow = tmpDayStartLine)
	openxlsx::writeData(wb, sheetTmp, x = "Mortality", startCol = 18, startRow = tmpDayStartLine + 1)
	openxlsx::writeData(wb, sheetTmp, x = "Comments", startCol = 19, startRow = tmpDayStartLine + 1)
	# 设置单元格格式
	openxlsx::addStyle(wb, sheetTmp, createStyle(fgFill = "#D3D3D3", border=c("top","right","bottom","left"), borderColour=c("black","black","black","black"), halign = "center", valign = "center"), rows = tmpDayStartLine, 
				cols = 1:19, gridExpand = FALSE, stack = TRUE)
	openxlsx::addStyle(wb, sheetTmp, createStyle(fgFill = "#D3D3D3", border=c("top","right","bottom","left"), borderColour=c("black","black","black","black"), halign = "center", valign = "center"), rows = tmpDayStartLine + 1, 
				cols = 1:19, gridExpand = FALSE, stack = TRUE)	
	# tmpSampleGroupData
	tmpSampleGroupData = data.frame(SampleGroupData, Cage = rep(NA, length(SampleGroupData$samples)))
	# body weight in day 'd'
	tmpDayBWdata = data.frame(Weight = as.numeric(BWmatrix[as.character(SampleGroupData$samples), tmpDay]), comments = rep(NA, length(SampleGroupData$samples)))
	# tumor volume in day 'd'
	tmpDayTVdata = data.frame(Length = as.numeric(TVlenMergeData[as.character(SampleGroupData$samples), tmpDay]), 
								Width = as.numeric(TVwidMergeData[as.character(SampleGroupData$samples), tmpDay]),
								Volume = as.numeric(TVmatrix[as.character(SampleGroupData$samples), tmpDay]), 
								comments = rep(NA, length(SampleGroupData$samples)))

	# 写入数据框
	openxlsx::writeData(wb, sheetTmp, tmpSampleGroupData, startCol = 1, startRow = tmpDayStartLine + 2, colNames = FALSE, rowNames = FALSE,
				borders = "columns", borderColour = "black")
	openxlsx::writeData(wb, sheetTmp, tmpDayBWdata, startCol = 4, startRow = tmpDayStartLine + 2, colNames = FALSE, rowNames = FALSE,
				borders = "surrounding", borderColour = "black")
	openxlsx::writeData(wb, sheetTmp, tmpDayTVdata, startCol = 6, startRow = tmpDayStartLine + 2, colNames = FALSE, rowNames = FALSE,
				borders = "surrounding", borderColour = "black")
	# 设置单元格格式
	# 右边框
	for(tmpCol in c(11,13,15,17,19)){
		openxlsx::addStyle(wb, sheetTmp, createStyle(border=c("right"), borderColour=c("black")), 
					rows = (tmpDayStartLine + 1):(tmpDayStartLine + 1 + length(SampleGroupData$samples)), 
					cols = tmpCol, gridExpand = FALSE, stack = TRUE)
	}
	# 下边框,合并单元格
	tmpDayDataStartLine = tmpDayStartLine + 2
	for(gvd in 1:length(TVdata)){
		openxlsx::addStyle(wb, sheetTmp, createStyle(border=c("bottom"), borderColour="black"), 
					rows = tmpDayDataStartLine + nrow(TVdata[[gvd]]$data) - 1, cols = 1:19, gridExpand = TRUE, stack = TRUE)
		openxlsx::mergeCells(wb, sheetTmp, cols=1, rows=tmpDayDataStartLine:(tmpDayDataStartLine + nrow(TVdata[[gvd]]$data) - 1))
		tmpDayDataStartLine = tmpDayDataStartLine + nrow(TVdata[[gvd]]$data)
	}
	# center
	startRowTmp = tmpDayStartLine + 2
	endRowTmp = tmpDayStartLine + 2 + nrow(tmpSampleGroupData) - 1
	for(gvc in 1:19){
		openxlsx::addStyle(wb, sheetTmp, createStyle(halign = "center", valign = "center"), 
				rows = startRowTmp:endRowTmp, cols = gvc, gridExpand = TRUE, stack = TRUE)
	}

}


openxlsx::saveWorkbook(wb, "analysis-result.xlsx")

file.remove(dir("." , pattern="(.png)$"))

