#数据整理部分
library("xlsx", lib.loc="~/R/win-library/3.2")
workdir = 'D:\\日常事务\\呼吸丙酮实验统计分析\\data'
alldata = as.data.frame(matrix(numeric(0),ncol=15))
colnames(alldata) = c('pNo','type','Age','Gender','Weight','Height','No','Acetone1','Acetone2','Acetone3','Acetone4','Acetone5','Acetone6','BGL','HBA')
index = 0
type = 0                                                   #类型
Acetone1 = NA
Acetone2 = NA
Acetone3 = NA
Acetone4 = NA
Acetone5 = NA
Acetone6 = NA
for (i in 1:5)
{
  ildata = read.xlsx(paste0(workdir,'\\2015 健康人.xls',collapse = NULL),colClasses = c(NA,NA,NA,'numeric','numeric','numeric','numeric','numeric',NA,'numeric','numeric','numeric','numeric','numeric','numeric'),sheetIndex = i)
  Gender = ildata[2,9]                                     #性别
  Age = ildata[2,10]                                       #年龄
  Weight = ildata[2,13]                                    #体重
  Height = ildata[2,14]                                    #身高
  pNo = i                                                  #参与者编号 整数
  plabel = 2
  while(TRUE)
  {
    if(ildata[plabel,3]==1)
    {
      No = as.numeric(substr(as.character(ildata[plabel,1]),3,4))   #访视编号  整数
      Acetone1 = ildata[plabel,4]                                   #呼吸丙酮1
      BGL = ildata[plabel,6]                                        #血糖
      HBA = ildata[plabel,8]                                        #丙酮
    }
    else if(ildata[plabel,3]==2)
    {
      Acetone2 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==3)
    {
      Acetone3 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==4)
    {
      Acetone4 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==5)
    {
      Acetone5 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==6)
    {
      Acetone6 = ildata[plabel,4]
    }
    plabel = plabel+1
    if(plabel>dim(ildata)[1])
    {
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      index = index+1
      alldata[index,]=row
      break
    }
    if(is.na(ildata[plabel,3]))
    {
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      index = index+1
      alldata[index,]=row
      break
    }
    if(ildata[plabel,3]==1)
    {
      
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      Acetone1 = NA
      Acetone2 = NA
      Acetone3 = NA
      Acetone4 = NA
      Acetone5 = NA
      Acetone6 = NA
      index = index+1
      alldata[index,]=row
    }
  }
}
type = 1                                                     #类型
for (i in 1:13)
{
  ildata = read.xlsx(paste0(workdir,'\\2015 糖尿病人.xls',collapse = NULL),colClasses = c(NA,NA,NA,'numeric','numeric','numeric','numeric','numeric',NA,'numeric','numeric','numeric','numeric','numeric','numeric'), sheetName = i)
  Gender = ildata[2,9]                                     #性别
  Age = ildata[2,10]                                       #年龄
  Weight = ildata[2,13]                                    #体重
  Height = ildata[2,14]                                    #身高
  pNo = i+5                                                #参与者编号 整数
  plabel = 2
  while(TRUE)
  {
    if(ildata[plabel,3]==1)
    {
      No = as.numeric(substr(as.character(ildata[plabel,1]),3,4))   #访视编号  整数
      Acetone1 = ildata[plabel,4]                                   #呼吸丙酮1
      BGL = ildata[plabel,6]                                        #血糖
      HBA = ildata[plabel,8]                                        #丙酮
    }
    else if(ildata[plabel,3]==2)
    {
      Acetone2 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==3)
    {
      Acetone3 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==4)
    {
      Acetone4 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==5)
    {
      Acetone5 = ildata[plabel,4]
    }
    else if(ildata[plabel,3]==6)
    {
      Acetone6 = ildata[plabel,4]
    }
    plabel = plabel+1
    if(plabel>dim(ildata)[1])
    {
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      index = index+1
      alldata[index,]=row
      break
    }
    if(is.na(ildata[plabel,3]))
    {
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      index = index+1
      alldata[index,]=row
      break
    }
    if(ildata[plabel,3]==1)
    {
      
      row = c(pNo,type,Age,Gender,Weight,Height,No,Acetone1,Acetone2,Acetone3,Acetone4,Acetone5,Acetone6,BGL,HBA)
      endace = 0
      Acetone1 = NA
      Acetone2 = NA
      Acetone3 = NA
      Acetone4 = NA
      Acetone5 = NA
      Acetone6 = NA
      index = index+1
      alldata[index,]=row
    }
  }
}

#数据分析部分

library(Daim)
alldata$bmi = alldata$Weight/((alldata$Height/100)^2)
r.ldh = roc(alldata[,8],alldata[,2],'1')
plot(r.ldh)



#正态性检验
tem = c()
for (i in c(1:dim(alldata)[1]))
{
  for (j in 8:13)
  {
    if(!is.na(alldata[i,j]))
    {
      tem = c(tem,alldata[i,j])
    }
  }
}
hist(tem,freq         = FALSE,ylim = c(0,0.4))
lines(density(tem),col='blue')
shapiro.test(tem)
qqnorm(tem, main="Q-Q plot: Price")
qqline(tem)

#取平均数
for (i in c(1:dim(alldata)[1]))
{
  num = 0
  total = 0
  for (j in 8:13)
  {
    if(!is.na(alldata[i,j]))
    {
      num = num+1
      total = total +alldata[i,j]
    }
  }
  alldata[i,'av_acetone'] = total/num
}


#取中位数
for (i in c(1:dim(alldata)[1]))
{
  num = 0
  tem = c()
  for (j in 8:13)
  {
    tem = c(tem,alldata[i,j])
  }
  alldata[i,'me_acetone'] = median(tem,na.rm = TRUE)
}

r.ldh = roc(alldata[,'me_acetone'],alldata[,'type'],'1')
plot(r.ldh)




bmig1 = (alldata[,'bmi']<18.5)&(alldata[,'av_acetone']<8)
bmig2 = (alldata[,'bmi']>=18.5)&(alldata[,'bmi']<24.99)
bmig3 = (alldata[,'bmi']>=25)&(alldata[,'bmi']<28)
bmig4 = (alldata[,'bmi']>=28)&(alldata[,'bmi']<32)
bmig5 = alldata[,'bmi']>=32

type0 = alldata[,'type']==0
type1 = alldata[,'type']==1


#分组ROC
bmig = bmig2
print(as.numeric(table(bmig&type0)['TRUE']))
print(as.numeric(table(bmig&type1)['TRUE']))
r.ldh = roc(alldata[bmig,'av_acetone'],alldata[bmig,'type'],'1')
plot(r.ldh)


library('e1071')

inputData <- data.frame(alldata[, c (16,20)], response = as.factor(alldata$type))
svmfit = svm(response ~., data = inputData, kernel = "polynomial", cost = 10, scale = FALSE) # linear svm, scaling turned OFF
print(svmfit)
plot(svmfit, inputData)
compareTable <- table (inputData$response, predict(svmfit))  # tabulate
mean(inputData$response != predict(svmfit)) # 19.44% misclassification error

