import xlrd, sys, os
"""
函数功能：将Excel表格中的数据转换成字典
参数列表：filePath-Excel文件路径，headerIndex-表头所在行的索引值，sheetIndex-表的索引
"""
def ExcelData2Dict(filePath, headerIndex=0, sheetIndex=0):
	try:
		excelDoc = xlrd.open_workbook(filePath)
	except:
		sys.exit
	table = excelDoc.sheet_by_index(sheetIndex)
	nrows = table.nrows #行数
	ncols = table.ncols #列数
	"""
	data = {
			'WorkID':{	#以workid作为索引
					'Speed':0,#速度
					'Command':2,#挖矿客户端类型
					'TimeLast':145,#在线时长
					'GPUInfo':'',  #GPU信息
					'CPUInfo':'',  #CPU信息
					'StartSource':'sysboot', #启动来源
					'Income':0,   #收入
					'ErrorMsg':'' #错误信息
				}
		}
	"""
	for rowIndex in range(headerIndex+1, nrows):
		dataTmp = table.row_values(rowIndex)
		infoTmp = {}
		infoTmp['Speed'] = int(float(dataTmp[1] or '-1'))
		infoTmp['Command'] = str(int(dataTmp[2] or '-1'))
		infoTmp['TimeLast'] = int(dataTmp[4] or '-1')
		infoTmp['GPUInfo'] = dataTmp[6]
		infoTmp['CPUInfo'] = dataTmp[7]
		infoTmp['StartSource'] = dataTmp[8]
		infoTmp['Income'] = int(dataTmp[9] or '-1')
		infoTmp['ErrorMsg'] = dataTmp[10]
		dataInfo[dataTmp[0]] = infoTmp
dataInfo = {}
strFilePath = r'D:\数据分析\共享赚宝\result.2017-12-13.xls'
if len(sys.argv) > 1:
	strFilePath = sys.argv[1]
ExcelData2Dict(strFilePath)
#分析
#无错误无收入的用户
noErrorIncomeData = {}
for key in dataInfo:
	if (dataInfo[key]['Income']<=0 and dataInfo[key]['ErrorMsg']==''):
		noErrorIncomeData[key] = dataInfo[key]
dataInfo_Count = len(dataInfo)
noErrorIncome_Count = len(noErrorIncomeData)
strOutput=[]
strLine = '总在线用户量为：' + str(dataInfo_Count)
print(strLine)
strOutput.append(strLine)
strLine = '无错误无收入的用户量为：%d, 占比总在线用户量%0.2f%%' % (noErrorIncome_Count, 100.0*noErrorIncome_Count/dataInfo_Count)
print(strLine)
strOutput.append(strLine)
del dataInfo	#删除多余的变量，节省内存
#运行时间低于5分钟的用户
strLine = '无错误无收入的用户中：'
print(strLine)
strOutput.append(strLine)
ThresholdTime  = 5 * 60
shortTimeData = {}
for key in noErrorIncomeData:
	if noErrorIncomeData[key]['TimeLast'] <= ThresholdTime:
		shortTimeData[key] = noErrorIncomeData[key]
shortTime_Count = len(shortTimeData)
strLine = '\t运行时间低于5分钟的用户量为：%d，占比%0.2f%%' % (shortTime_Count, 100.0*shortTime_Count/noErrorIncome_Count)
print(strLine)
strOutput.append(strLine)
#速度为0的用户
zeroSpeed_UserCount = 0
for key in noErrorIncomeData:
	if noErrorIncomeData[key]['Speed'] <= 0:
		zeroSpeed_UserCount += 1
strLine = '\t速度为0的用户量为：%d，占比%0.2f%%' % (zeroSpeed_UserCount, 100.0*zeroSpeed_UserCount/noErrorIncome_Count)
print(strLine)
strOutput.append(strLine)
#运行时间低于5分钟，且速度不为0的用户
strLine = '运行时间低于5分钟的用户中：'
print(strLine)
strOutput.append(strLine)
shortTimeNoZeroSpeed_UserCount = 0
for key in shortTimeData:
	if shortTimeData[key]['Speed'] > 0:
		shortTimeNoZeroSpeed_UserCount += 1
strLine = '\t速度不为0的用户量：%d，占比%0.2f%%' % (shortTimeNoZeroSpeed_UserCount,100.0*shortTimeNoZeroSpeed_UserCount/shortTime_Count)
print(strLine)
strOutput.append(strLine)
#运行时间低于5分钟，且速度为0的用户
shortTimeZeroSpeedData = {}
for key in shortTimeData:
	if shortTimeData[key]['Speed'] <= 0:
		shortTimeZeroSpeedData[key] = shortTimeData[key]
shortTimeZeroSpeed_UserCount = len(shortTimeZeroSpeedData)
strLine = '\t速度为0的用户量：%d，占比%0.2f%%' % (shortTimeZeroSpeed_UserCount, 100.0*shortTimeZeroSpeed_UserCount/shortTime_Count)
print(strLine)
strOutput.append(strLine)
#运行时间低于5分钟，且速度为0的用户，挖矿种类
MinerType = {}
for key in shortTimeZeroSpeedData:
	typeTmp = shortTimeZeroSpeedData[key]['Command']
	if typeTmp in MinerType.keys(): #判断字典中是否存在某个键
		MinerType[typeTmp] += 1
	else:
		MinerType[typeTmp] = 1
strLine = '运行时间低于5分钟，且速度为0的用户中：'
print(strLine)
strOutput.append(strLine)
minerFlag = {'1' : 'ETC        ',
			 '2' : 'N卡ZCash   ',
			 '3' : 'A卡ZCash   ',
			 '4' : 'N卡XMR_64位',
			 '5' : 'A卡XMR_64位',
			 '6' : 'A卡XMR_32位',
			 '7' : 'CPU_XMR    ',}
MinerType = sorted(MinerType.items(), key=lambda item:item[1], reverse=True)	#sorted函数按value值对字典排序。注意排序后的返回值是一个list，而原字典中的键值对被转换为了list中的元组
for key in MinerType:
	#print('\t挖 %s 的用户有 %d 个，占比%0.2f%%' % (minerFlag[key], MinerType[key], 100.0*MinerType[key]/shortTimeZeroSpeed_UserCount))	
	strLine = '\t挖 %s 的用户有 %d 个，占比%0.2f%%' % (minerFlag[key[0]], key[1], 100.0*key[1]/shortTimeZeroSpeed_UserCount)
	print(strLine)
	strOutput.append(strLine)
#将分析结果写入文本
strSaveDir = os.path.splitext(strFilePath)[0]
if not os.path.exists(strSaveDir):
    os.mkdir(strSaveDir)
strSavePath = os.path.join(strSaveDir, '无错误无收入.txt')
try:
	f = open(strSavePath, 'w+')
	for strTmp in strOutput:
		f.write(strTmp + '\n')
	f.close()
finally:
	if 'f' in locals():
		f.close()
