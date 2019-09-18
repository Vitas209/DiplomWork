import xlrd as xl
def minm(S):
    Mass=[]
    for i in range(sheet.nrows):
        r=sheet.row_values(i)
        Mass.append(r)
    K=sum(Mass[0])-1
    while K>0:
        Pyt=[0]
        Pyt1=[]
        St=''
        Arr=[]
        a=K
        A1=0
        for i in range(sheet.nrows-1):
            if i%2==0:
                Pyt1.append(str(round(i/2)+1))
                Pyt.append(float('inf'))
        for i in range(round(sheet.nrows/2)):
            for j in range(round(sheet.nrows/2)):
                if (int(Mass[i*2][j+1])>0) and (Pyt[j]>Mass[i*2+1][j+1]+Pyt[i]):
                    Pyt[j]=Mass[i*2+1][j+1]+Pyt[i]
                    Pyt1[j]=Pyt1[i]+' '+str(j+1)
        St=Pyt1[-1]
        St=St+' '
        while len(St)!=0:
            A1=St.find(' ',0,len(St))
            Arr.append(int(St[0:A1]))
            St=St[A1+1:len(St)]
        for i in range(len(Arr)-1):
            if a>Mass[(Arr[i])*2-2][(Arr[i+1])]:
                a=Mass[(Arr[i])*2-2][(Arr[i+1])]
        for i in range(len(Arr)-1): 
            Mass[(Arr[i])*2-2][(Arr[i+1])]=Mass[(Arr[i])*2-2][(Arr[i+1])]-a
        С=Pyt[-2]*a
        S=S+Pyt[-2]*a
        K=K-a
        print('По маршруту ',Pyt1[-1]," Нужно отправить",a," ед. груза, чтобы затраты были минимальны и составили ",С," у.е")
    return S
rd=xl.open_workbook('data.xlsx',formatting_info=False)
sheet=rd.sheet_by_index(0)
S=0
S=minm(S)
print('Общие затраты составляют ',S," у.е.")