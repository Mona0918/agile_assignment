import pandas as pd
list1=[]
list_count={}
team_user=[]
flag=0
dict={}
df1=pd.read_excel(io='C:/Users/Mona/PycharmProjects/agile/Input sheets/input_sheet1.xlsx',
               dtype={'Name':str,'Team Name':str,'User ID':int})
df2=pd.read_excel(io='C:/Users/Mona/PycharmProjects/agile/Input sheets/input_sheet2.xlsx',
               dtype={'Name':str,'User ID':int,'total_statements':int,'total_reasons':int})
df_list=list(df1.loc[:,"Team Name"])
df_ID=list(df1.loc[:,"User ID"])
list1=list(set(df_list))
for j in range(len(list1)):
    for k in range(len(df_list)):
        if list1[j] == df_list[k]:
            flag=1
            team_user=team_user+[df_ID[k]]
        else:
            flag=0
    dict[list1[j]]=team_user
    team_user=[]
for i in list1:
    list_count[i]=df_list.count(i)
df_list1=list(df2.loc[:,"User ID"])
df_stmt=list(df2.loc[:,"total_statements"])
df_reason=list(df2.loc[:,"total_reasons"])
ID_stmt={}
ID_reason={}
for k in range(len(df_list1)):
    ID_stmt[df_list1[k]]=df_stmt[k]
    ID_reason[df_list1[k]]=df_reason[k]
key1=list(ID_stmt.keys())
key2=list(ID_reason.keys())
sum_stmt=0
avg_stmt=0
sum_reason=0
avg_reason=0
team_avgstmt={}
team_avgreason={}
for i in dict:
    for j in dict[i]:
        for k in key1:
            if j == k:
                sum_stmt=sum_stmt+ID_stmt[j]
        avg_stmt=sum_stmt/list_count[i]
    team_avgstmt[i]=avg_stmt
    sum_stmt=0
    avg_stmt=0
for i in dict:
    for j in dict[i]:
        for k in key2:
            if j == k:
                sum_reason=sum_reason+ID_reason[j]
        avg_reason=sum_reason/list_count[i]
    team_avgreason[i]=avg_reason
    sum_reason=0
    avg_reason=0
stmt_reason={}
for i in list1:
    sum=team_avgstmt[i]+team_avgreason[i]
    stmt_reason[i]=[team_avgstmt[i]]+[team_avgreason[i]]+[sum]
sorted_stmt_reason=sorted(stmt_reason.items(),reverse=True)
ky=[]
for i in sorted_stmt_reason:
    for j in range(len(i)):
        if j==0:
            ky=ky+[i[j]]
s=[]
r=[]
for i in sorted_stmt_reason:
    for j in range(len(i)):
        if j==1:
            for k in range(len(i[j])):
                if k == 0:
                    s=s+[i[j][k]]
                elif k == 1:
                    r=r+[i[j][k]]
Rank=[1,2,3,4,5,6,7,8]
Leaderboard=ky
stmt=s
reason=r
columns=['Team Rank','Thinking Teams Leaderboard','Average_Statements','Average_Reasons']
df=pd.DataFrame(list(zip(Rank,Leaderboard,stmt,reason)),columns=columns)
df.to_excel("C:/Users/Mona/PycharmProjects/agile/Output sheets/Leaderboard.xlsx",sheet_name="LeaderBoard Teamwise(Output)",columns=columns)