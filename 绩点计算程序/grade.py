#����py3+pandas
#���޸ĳɼ����λ�ã���ȨȨ��
from pandas import read_excel,ExcelFile,Series
mygrade=read_excel('D:/Desktop/grade.xlsx')#����ɼ��洢λ��
gradelevel=['A+','A','B+','B','C+','C','D+','D','F']
gradescore=[4.5,4,3.5,3,2.5,2,1.5,1,0]
#��ԭ�ɼ���ͳ��ѧλ�β�����degree_course�б���
row=mygrade.shape[0]
degree_course=[]
for i in range(0,row):
    if(mygrade.iloc[i,14]=='��'):
        degree_course.append(mygrade.iloc[i,3])
gradeweight=[1.2,1,1]#���Ƽ���Ȩ��,�ֱ�Ϊѧλ�Ρ�ѡ�޿Ρ���������Ȩ������㷽ʽ��1.2��1��1
excel = ExcelFile('D:/Desktop/grade1.xlsx')#�����˳ɼ�����λ��
names=excel.sheet_names
data={}
s=Series(range(0,len(names)),index=names,dtype='float64')
for j in range(0,len(names)):
    data[names[j]]=read_excel(excel,names[j])
    col=data[names[j]].shape[1]
    row=data[names[j]].shape[0]
    #���Ƽ���Ȩ��
    for i in range(0,row):
        if(degree_course.count(data[names[j]].iloc[i,3])):
            data[names[j]].iloc[i,5]=gradeweight[0]
        elif(data[names[j]].iloc[i,4]=='����ѡ�޿�'):
            data[names[j]].iloc[i,5]=gradeweight[1]
        else:
            data[names[j]].iloc[i,5]=gradeweight[2]
    #�����Ȩ����
    mem=0
    demo=0
    for i in range(0,row):
        mem+=data[names[j]].iloc[i,5]*data[names[j]].iloc[i,6]*data[names[j]].iloc[i,7]
        demo+=data[names[j]].iloc[i,5]*data[names[j]].iloc[i,6]
    result=mem/demo
    s[names[j]]=result
s=s.sort_values(ascending=False)
print(s)#�������������˳�򽫱�����������