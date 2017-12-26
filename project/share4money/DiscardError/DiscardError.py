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
		infoTmp['ErrorMsg'] = str(dataTmp[10])
		dataInfo[dataTmp[0]] = infoTmp
dataInfo = {}
strFilePath = r'C:\Users\Suo\Desktop\test.xls' #r'D:\数据分析\共享赚宝\result.2017-12-13.xls'
if len(sys.argv) > 1:
	strFilePath = sys.argv[1]
ExcelData2Dict(strFilePath)
strOutput=[]
#获取GPU类型、[[型号、大小]、]  0--Intel、1--AMD、2--Nvidia
def GetGPUInfo(strInfo):
    detailList = []
    iTypeRet = 0
    iIntel, iAmd, iNvidia = 0, 0, 0
    listTmp = strInfo.split('_')
    for item in listTmp:
        if item[0] == '0':
            iIntel += 1
        elif item[0] == '1':
            iAmd += 1
        elif item[0] == '2':
            iNvidia += 1
    #根据分析的结果进行判断，并将相应的信息保存
    if iNvidia > 0:
        iTypeRet = 2
        for item in listTmp:
            if item[0] == '2':
                singleDetailList = item.split('|')
                detailList.append(singleDetailList[1:])
    elif iAmd > 0:
        iTypeRet = 1
        for item in listTmp:
            if item[0] == '1':
                singleDetailList = item.split('|')
                detailList.append(singleDetailList[1:])
    elif iIntel > 0: #最后判断是Intel集显
        for item in listTmp:
            if item[0] == '0':
                singleDetailList = item.split('|')
                detailList.append(singleDetailList[1:])
    return iTypeRet, detailList
#分析有错误信息，有显卡信息，无收入的用户
'''[(errorMsg, MinerType, [GPUType]), (,)]'''
NvidiaTargetList, AmdTargetList = [], []
targetCount = 0
for key in dataInfo:
    if dataInfo[key]['ErrorMsg'] != '' and dataInfo[key]['GPUInfo'] != '' and dataInfo[key]['Income']<= 0:
        targetCount += 1
        iType, detailList = GetGPUInfo(dataInfo[key]['GPUInfo'])
        if iType == 1 and len(detailList) == 1: #只分析单显卡
            AmdTargetList.append((dataInfo[key]['ErrorMsg'], dataInfo[key]['Command'], detailList[0]))
        elif iType == 2 and len(detailList) == 1:
            NvidiaTargetList.append((dataInfo[key]['ErrorMsg'], dataInfo[key]['Command'], detailList[0]))
#输出总体信息
totalInfoCount = len(dataInfo)
strTmp = '总数据量为：{:<5}，其中有错误信息、有显卡信息、无收入的用户量为：{:<5}，占比{:.2f}%'.format(totalInfoCount, targetCount, targetCount * 100.0 / totalInfoCount)
print(strTmp)
strOutput.append(strTmp)
strTmp = '1：Nvidia显卡用户量为：{:<5}，占比{:.2f}%'.format(len(NvidiaTargetList), len(NvidiaTargetList) * 100.0 / targetCount)
print(strTmp)
strOutput.append(strTmp)
strTmp = '   AMD显卡用户量为：{:<5}，占比{:.2f}%'.format(len(AmdTargetList), len(AmdTargetList) * 100.0 / targetCount)
print(strTmp)
strOutput.append(strTmp)
#分析错误信息#'''[(MinerType, [GPUType,,]), (,)]'''
NvidiaVerErrorList, AmdVerErrorList = [], []
iProcessExit = 0
def CalcErrorInfo(listTmp, desVerErrList):
    global iProcessExit
    for item in listTmp:
        if 'insufficient' in item[0] or 'cannot load nvml' in item[0] or 'no cuda device' in item[0] or 'no amd device' in item[0]:
            desVerErrList.append(item[1:])
        elif item[0] == '1.0' or item[0] == '2.0':
            iProcessExit += 1
CalcErrorInfo(NvidiaTargetList, NvidiaVerErrorList)
CalcErrorInfo(AmdTargetList, AmdVerErrorList)
iDriveVersionError = len(NvidiaVerErrorList) + len(AmdVerErrorList)
strTmp = '2：驱动程序不匹配的量为：{:<5}，占比{:.2f}%'.format(iDriveVersionError, iDriveVersionError * 100.0 / targetCount)
print(strTmp)
strOutput.append(strTmp)
strTmp = '   挖矿客户端进程退出的量为：{:<5}，占比{:.2f}%'.format(iProcessExit, iProcessExit * 100.0 / targetCount)
print(strTmp)
strOutput.append(strTmp)
#删除原始数据，释放内存
del dataInfo

#按挖矿客户端的Top10错误
minerFlag = {'1' : 'ETC        ',
			 '2' : 'N卡ZCash   ',
			 '3' : 'A卡ZCash   ',
			 '4' : 'N卡XMR_64位',
			 '5' : 'A卡XMR_64位',
			 '6' : 'A卡XMR_32位',
			 '7' : 'CPU_XMR    ',}
#按矿种分析  
def GetInfoAccordingMinerType(verErrorList):
    for key in minerFlag:
        discardDict = {}
        iErrorCount = 0
        for item in verErrorList:
            if item[0] == key:
                strType = item[1][0]
                iErrorCount += 1
                if strType in discardDict.keys():
                    discardDict[strType] += 1
                else:
                    discardDict[strType] = 1
        if iErrorCount > 0:
            discardDict = sorted(discardDict.items(), key = lambda item: item[1], reverse = True)
            strTmp = '{}错误总数为{:>5}，占比{:.2f}%'.format(minerFlag[key], iErrorCount, iErrorCount*100.0/len(verErrorList))
            print(strTmp)
            strOutput.append(strTmp)
            for item in discardDict:
                if item[1] >= 10:
                    strTmp = '\t型号：{:>30}，总量：{:>5}，占比{:.2f}%'.format(item[0], item[1], item[1]*100.0/iErrorCount)
                    print(strTmp)
                    strOutput.append(strTmp)
if len(NvidiaVerErrorList) > 0:
    strTmp = 'Nvidia单显卡驱动程序不匹配的分析结果：总量：{:>5}，占比{:.2f}%'.format(len(NvidiaVerErrorList), len(NvidiaVerErrorList)*100.0/iDriveVersionError)
    print(strTmp)
    strOutput.append(strTmp)
    GetInfoAccordingMinerType(NvidiaVerErrorList)
if len(AmdVerErrorList) > 0:
    strTmp = 'AMD单显卡驱动程序不匹配的分析结果：总量：{:>5}，占比{:.2f}%'.format(len(AmdVerErrorList), len(AmdVerErrorList)*100.0/iDriveVersionError)
    print(strTmp)
    strOutput.append(strTmp)
    GetInfoAccordingMinerType(AmdVerErrorList)

#将分析结果写入文本
strSaveDir = os.path.splitext(strFilePath)[0]
if not os.path.exists(strSaveDir):
    os.mkdir(strSaveDir)
strSavePath = os.path.join(strSaveDir, '显卡驱动不匹配.txt')
try:
	f = open(strSavePath, 'w+')
	for strTmp in strOutput:
		f.write(strTmp + '\n')
	f.close()
finally:
	if 'f' in locals():
		f.close()

