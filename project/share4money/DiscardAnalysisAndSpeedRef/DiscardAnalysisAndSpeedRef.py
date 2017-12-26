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
strOutput=[]
#分析
#首先分析显卡占比信息，并且把相关的信息存表
nvidiaCount, amdCount, intelCount = 0, 0, 0 #显卡分类计量
mutiNvidiaCount, mutiAmdCount = 0, 0 #多块显卡
'''{WorkID键:{'Speed':,'Command':,'GPUType':,'GPUSize':}}'''
NvidiaTragetDict, AmdTragetDict = {}, {}    #非零的显卡信息,我们分析的重点
'''[(GPUType, GPUSize), (,)]'''
NvidiaTotalDict, AmdTotalDict = [], [] #总的显卡信息
for key in dataInfo:
    strGPUInfo = dataInfo[key]['GPUInfo']
    iIntel, iAmd, iNvidia = 0, 0, 0
    listTmp = []
    if strGPUInfo != '':
        listTmp = strGPUInfo.split('_')
        for item in listTmp:
            if item[0] == '0':
                iIntel += 1
            elif item[0] == '1':
                iAmd += 1
            elif item[0] == '2':
                iNvidia += 1
    #根据分析的结果进行判断，并将相应的信息保存
    if iNvidia > 0:
        nvidiaCount += 1
        if iNvidia > 1: #多块Nvidia显卡
            mutiNvidiaCount += 1
            continue
        typeTmp, sizeTmp = '', ''
        for item in listTmp:
            if item[0] == '2':
                detailList = item.split('|')
                typeTmp = detailList[1]
                sizeTmp = detailList[2]
                break
        NvidiaTotalDict.append((typeTmp, sizeTmp))
        #只分析速度不为0的单独立显卡
        if dataInfo[key]['Speed'] > 0:
            NvidiaTragetDict[key]={}
            NvidiaTragetDict[key]['Speed'] = dataInfo[key]['Speed']
            NvidiaTragetDict[key]['Command'] = dataInfo[key]['Command']
            NvidiaTragetDict[key]['GPUType'] = typeTmp
            NvidiaTragetDict[key]['GPUSize'] = sizeTmp
    elif iAmd > 0:
        amdCount += 1
        if iAmd > 1: #多块Amd显卡
            mutiAmdCount += 1
            continue
        typeTmp, sizeTmp = '', ''
        for item in listTmp:
            if item[0] == '1':
                detailList = item.split('|')
                typeTmp = detailList[1]
                sizeTmp = detailList[2]
                break
        AmdTotalDict.append((typeTmp, sizeTmp))
        #只分析速度不为0的单独立显卡
        if dataInfo[key]['Speed'] > 0:
            AmdTragetDict[key]={}
            AmdTragetDict[key]['Speed'] = dataInfo[key]['Speed']
            AmdTragetDict[key]['Command'] = dataInfo[key]['Command']
            AmdTragetDict[key]['GPUType'] = typeTmp
            AmdTragetDict[key]['GPUSize'] = sizeTmp
    elif iIntel > 0: #最后判断是Intel集显
        intelCount += 1
totalInfoCount = len(dataInfo)
succInfoCount = nvidiaCount + amdCount + intelCount
strTmp = '总数据量为：{:<5}，成功获取显卡信息量为：{:<5}，占比{:.2f}%'.format(totalInfoCount, succInfoCount, succInfoCount * 100.0 / totalInfoCount)
print(strTmp)
strOutput.append(strTmp)
strTmp = '以下数据以成功获取的显卡信息为基础分析！'
print(strTmp)
strOutput.append(strTmp)
#获取显卡总体信息
def GetDiscardTotalInfo(totalCard, multiCard, targetCard):
    strTmp = '\t总量为：{:<4}，占比{:.2f}%；多显卡量为{}，占比该显卡总量的{:.2f}%'.format(totalCard, totalCard * 100.0 / succInfoCount, multiCard, multiCard * 100.0 / totalCard)
    print(strTmp)
    strOutput.append(strTmp)
    strTmp = '\t速度不为0的显卡量为{:<4}，占比该显卡总量的{:.2f}%'.format(targetCard, targetCard * 100.0 / totalCard)
    print(strTmp)
    strOutput.append(strTmp)
strTmp = 'Nvidia显卡整体信息：'
print(strTmp)
strOutput.append(strTmp)
GetDiscardTotalInfo(nvidiaCount, mutiNvidiaCount, len(NvidiaTragetDict))
strTmp = 'Amd显卡整体信息：'
print(strTmp)
strOutput.append(strTmp)
GetDiscardTotalInfo(amdCount, mutiAmdCount, len(AmdTragetDict))
strTmp = 'Intel显卡整体信息：'
print(strTmp)
strOutput.append(strTmp)
strTmp = '\t总量为：{:<4}，占比{:.2f}%'.format(intelCount, intelCount * 100.0 / succInfoCount)
print(strTmp)
strOutput.append(strTmp)

#删除原始数据，释放内存
del dataInfo
#最少显卡数量为10
DISCARD_MINCOUNT = 10
#分析一下两种显卡的型号占比:单独立显卡，包括0速度
def GetSingleDiscardPercent(totalDict):
    TypeAN = {}
    for item in totalDict:
        if item[0] in TypeAN.keys():
            TypeAN[item[0]] += 1
        else:
            TypeAN[item[0]] = 1
    TypeAN= sorted(TypeAN.items(), key = lambda item: item[1], reverse = True)
    for item in TypeAN:
        if item[1] >= 10:
            strTmp = '\t型号：{:<30},数量{:<4},占比{:.2f}%'.format(item[0], item[1], item[1]*100.0/len(totalDict))
            print(strTmp)
            strOutput.append(strTmp)
strTmp = 'Nvidia显卡（总量为{}）各型号占比：单独立显卡，包括0速度'.format(len(NvidiaTotalDict))
print(strTmp)
strOutput.append(strTmp)
GetSingleDiscardPercent(NvidiaTotalDict)
strTmp = 'AMD显卡（总量为{}）各型号占比：单独立显卡，包括0速度'.format(len(AmdTotalDict))
print(strTmp)
strOutput.append(strTmp)
GetSingleDiscardPercent(AmdTotalDict)
#删除原始数据，释放内存
del AmdTotalDict
del NvidiaTotalDict
#只分析速度不为0的
#先分析Nvidia显卡
strTmp = '只分析速度不为0的单独立显卡'
print(strTmp)
strOutput.append(strTmp)
minerFlag = {'1' : 'ETC        ',
			 '2' : 'N卡ZCash   ',
			 '3' : 'A卡ZCash   ',
			 '4' : 'N卡XMR_64位',
			 '5' : 'A卡XMR_64位',
			 '6' : 'A卡XMR_32位',
			 '7' : 'CPU_XMR    ',}
def GetSpeedFlag(listTmp):
    listTmp.sort()
    index = len(listTmp) // 3
    return listTmp[0], listTmp[-1], listTmp[index:-index]
#按矿种分析  
def GetInfoAccordingMinerTypeDict(MinerDict):
    MinerInfo = {}  #{矿种:{显卡类型:[(速度，显存大小)]}}
    for key in MinerDict:
        minerType = MinerDict[key]['Command']
        if minerType in MinerInfo:
            if not MinerDict[key]['GPUType'] in MinerInfo[minerType].keys():
                MinerInfo[minerType][MinerDict[key]['GPUType']] = []
        else:
            MinerInfo[minerType] = {}
            MinerInfo[minerType][MinerDict[key]['GPUType']] = []
        MinerInfo[minerType][MinerDict[key]['GPUType']].append((MinerDict[key]['Speed'], MinerDict[key]['GPUSize']))
    #将矿种按显卡数量排序，数量最少为10
    for key in MinerInfo:
        subMinerInfo = MinerInfo[key]
        subMinerInfo = sorted(subMinerInfo.items(), key = lambda item:len(item[1]), reverse=True)
        strTmp = minerFlag[key]
        print(strTmp)
        strOutput.append(strTmp)
        for item in subMinerInfo:
            if len(item[1]) >= DISCARD_MINCOUNT:
                speedList = []
                for subItem in item[1]:
                    speedList.append(subItem[0])
                minSpeed, maxSpeed, referSpeedList = GetSpeedFlag(speedList)
                strTmp = '\t型号：{:<30}，数量：{:<4},占比{:.2f}%，最低速度{}，最高速度{}，参考速度区间为{}~{}'.format(item[0], len(item[1]), len(item[1])*100.0/len(MinerDict), minSpeed, maxSpeed, referSpeedList[0], referSpeedList[-1])
                print(strTmp)
                strOutput.append(strTmp)
                
strTmp = 'Nvidia单显卡（总量为{}）详细信息：速度不为0'.format(len(NvidiaTragetDict))
print(strTmp)
strOutput.append(strTmp)
GetInfoAccordingMinerTypeDict(NvidiaTragetDict)
strTmp = 'AMD单显卡（总量为{}）详细信息：速度不为0'.format(len(AmdTragetDict))
print(strTmp)
strOutput.append(strTmp)
GetInfoAccordingMinerTypeDict(AmdTragetDict)

#将分析结果写入文本
strSaveDir = os.path.splitext(strFilePath)[0]
if not os.path.exists(strSaveDir):
    os.mkdir(strSaveDir)
strSavePath = os.path.join(strSaveDir, '显卡信息与速度参考.txt')
try:
	f = open(strSavePath, 'w+')
	for strTmp in strOutput:
		f.write(strTmp + '\n')
	f.close()
finally:
	if 'f' in locals():
		f.close()


