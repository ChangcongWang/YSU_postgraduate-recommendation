#环境py3+pandas
#请修改成绩存放位置，加权权重
from pandas import read_excel,ExcelFile,Series
mygrade=read_excel('D:/Desktop/grade.xlsx')#自身成绩存储位置
gradelevel=['A+','A','B+','B','C+','C','D+','D','F']
gradescore=[4.5,4,3.5,3,2.5,2,1.5,1,0]
#按原成绩单统计学位课并存在degree_course列表中
row=mygrade.shape[0]
degree_course=[]
for i in range(0,row):
    if(mygrade.iloc[i,14]=='是'):
        degree_course.append(mygrade.iloc[i,3])
gradeweight=[1.2,1,1]#各科绩点权重,分别为学位课、选修课、其他，加权绩点计算方式：1.2、1、1
excel = ExcelFile('D:/Desktop/grade1.xlsx')#所有人成绩储存位置
names=excel.sheet_names
data={}
s=Series(range(0,len(names)),index=names,dtype='float64')
for j in range(0,len(names)):
    data[names[j]]=read_excel(excel,names[j])
    col=data[names[j]].shape[1]
    row=data[names[j]].shape[0]
    #各科绩点权重
    for i in range(0,row):
        if(degree_course.count(data[names[j]].iloc[i,3])):
            data[names[j]].iloc[i,5]=gradeweight[0]
        elif(data[names[j]].iloc[i,4]=='公共选修课'):
            data[names[j]].iloc[i,5]=gradeweight[1]
        else:
            data[names[j]].iloc[i,5]=gradeweight[2]
    #计算加权绩点
    mem=0
    demo=0
    for i in range(0,row):
        mem+=data[names[j]].iloc[i,5]*data[names[j]].iloc[i,6]*data[names[j]].iloc[i,7]
        demo+=data[names[j]].iloc[i,5]*data[names[j]].iloc[i,6]
    result=mem/demo
    s[names[j]]=result
s=s.sort_values(ascending=False)
print(s)#最终输出按绩点顺序将表名进行排列