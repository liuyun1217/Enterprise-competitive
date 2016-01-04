##============================================================================================================##
##测试部分
indusData <- data1[grep(pattern = "道路运输业",x = data1$所属证监会行业名称...行业级别..大类行业,fixed = TRUE),]
DaoLu2008 <- indusData[grep(indusData$年份 = 2008,fixed = TRUE),]
DaoLu2007 <- indusData[grep(pattern = 2007,indusData$年份,fixed = TRUE),]
MeanDaoLu2007 <- mean(DaoLu2007,na.rm = TRUE)
StdDaoLu2007 <- sd(as.numeric(DaoLu2007$资产负债率.),na.rm = TRUE)
ShengYuFuZhaiDaoLu2007 <- (DaoLu2007$资产负债率. - MeanDaoLu2007)/StdDaoLu2007
DaoLu2007$剩余负债能力 <- ShengYuFuZhaiDaoLu2007

NameUniq <- unique(data1$所属证监会行业名称...行业级别..大类行业)
YearsUniq <- 2007:2014
newdcl <- data.frame(dcl2[1,],stringsAsFactors = FALSE)
newdcl <- newdcl[-1,]
length(NameUniq)
##============================================================================================================##
##主程序部分

##打开表格
data1 <- data.frame(read.xlsx2("v3.0（企业市场竞争力、现金持有、剩余负债能力）1.xlsx",sheetIndex = 1,stringsAsFactors=FALSE))
##创建一个空的表格，与打开的表格格式一样，表头一样
res1 <- data.frame(data1[1,],stringsAsFactors = FALSE)
res1 <- res1[-1,]

##读取行业名称和年份
NameUniq <- unique(data1$所属证监会行业名称...行业级别..大类行业)
YearsUniq <- 2007:2014
##开始处理，首先按照行业名称一个一个筛选
for(iname in 1:(length(NameUniq)-1)){
    ##读取本次处理的行业名称所有的数据(包括每一年)
    NameData <- data1[grep(pattern = NameUniq[iname],x = data1$所属证监会行业名称...行业级别..大类行业,fixed = TRUE),]
    ##按年份进行第二次筛选
    for(iyear in 2007:2014){
        ##传进来的某行业名称所有数据的某一年数据
        YearNameData <- NameData[grep(pattern = iyear,x = NameData$年份,fixed = TRUE),]
        ##计算平均值，标准差，并且计算标准化写入到剩余负债能力表格里
        MeanTempZiChanData <- mean(as.numeric(YearNameData$资产负债率.),na.rm = TRUE)
        StdTempZiChanData <- sd(as.numeric(YearNameData$资产负债率.),na.rm = TRUE)
        ShengYuFuZhaiData <- (as.numeric(YearNameData$资产负债率. )- MeanTempZiChanData)/StdTempZiChanData
        YearNameData$剩余负债能力 <- ShengYuFuZhaiData
        ##计算平均值，标准差，并且计算标准化写入到现金持有水平表格里
        MeanTempXianJinData <- mean(as.numeric(YearNameData$现金比率),na.rm = TRUE)
        StdTempXianJinData <- sd(as.numeric(YearNameData$现金比率),na.rm = TRUE)
        XianJinChiYouData <- (as.numeric(YearNameData$现金比率) - MeanTempXianJinData)/StdTempXianJinData
        YearNameData$现金持有水平 <- XianJinChiYouData
        ##计算平均值，并且计算企业竞争力写入表格里
        MeanTempXiaoShouData <- mean(as.numeric(YearNameData$销售增长.Growth.),na.rm = TRUE)
        QiYe <- as.numeric(YearNameData$销售增长.Growth.) - MeanTempXiaoShouData
        YearNameData$企业市场竞争力 <- QiYe
        ##将此次计算的某行业某年份的填好待计算数据的表格合并到结果表格res1里
        res1 <- rbind(res1,YearNameData)
    }
}
##写入到excel表格里
write.xlsx2(res1,"res1.xlsx")