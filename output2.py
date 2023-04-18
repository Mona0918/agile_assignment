import pandas as pd
df1=pd.read_excel(io='C:/Users/Mona/PycharmProjects/agile/Input sheets/input_sheet1.xlsx',
               dtype={'Name':str,'Team Name':str,'User ID':int})
df2=pd.read_excel(io='C:/Users/Mona/PycharmProjects/agile/Input sheets/input_sheet2.xlsx',
               dtype={'Name':str,'User ID':int,'total_statements':int,'total_reasons':int})
name_ID={}
ID_stmt_reas={}
ID_stmt_reas1={}
df1_name=list(df1.loc[:,"Name"])
df1_ID=list(df1.loc[:,"User ID"])
name_ID2={}
for i in range(len(df1_ID)):
    name_ID[df1_ID[i]]=df1_name[i]
    name_ID2[df1_name[i]]=df1_ID[i]
df2_ID=list(df2.loc[:,"User ID"])
df2_stmt=list(df2.loc[:,"total_statements"])
df2_reason=list(df2.loc[:,"total_reasons"])
for i in range(len(df2_ID)):
    ID_stmt_reas[df2_ID[i]]=[df2_stmt[i]]+[df2_reason[i]]
    ID_stmt_reas1[df2_ID[i]]=df2_stmt[i]+df2_reason[i]
sorted_ID_stmt_reas=sorted(ID_stmt_reas1.items(),key=lambda x:(x[1]*-1,x[0]))
sorted_name_ID={}
k=list(name_ID.keys())
for i in sorted_ID_stmt_reas:
    for j in range(len(i)):
            for r in k:
                if i[0] == r:
                    sorted_name_ID[r]=name_ID[r]
rank=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21]
name=[]
stmt=[]
reas=[]
for i in sorted_name_ID:
    name=name+[sorted_name_ID[i]]
sor={}
name_total={}
t=[]
for i in sorted_ID_stmt_reas:
    t=t+[i[1]]
for j in range(len(name)):
        name_total[name[j]]=t[j]
s=sorted(name_total.items(),key=lambda x:(x[1]*-1,x[0]))
name1=[]
ID1=[]
for i in s:
    name1=name1+[i[0]]
sorted_name_ID1={}
for i in s:
    for r in name1:
        if i[0] == r:
            sorted_name_ID1[r]=name_ID2[r]
for i in sorted_name_ID1:
    ID1=ID1+[sorted_name_ID1[i]]
for i in ID1:
    sr=ID_stmt_reas[i]
    sor[i]=sr
for i in sor:
    for j in range(len(sor[i])):
        if j == 0:
            stmt=stmt+[sor[i][j]]
        else:
            reas=reas+[sor[i][j]]
columns=['Rank','Name','UID','No.of Statements','No.of Reasons']
df=pd.DataFrame(list(zip(rank,name1,ID1,stmt,reas)),columns=columns)
df.to_excel("C:/Users/Mona/PycharmProjects/agile/Output sheets/Leaderboard_Individual.xlsx",sheet_name="LeaderBoard Individual(Output)",columns=columns)