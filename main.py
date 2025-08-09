import pandas as pd
import os

File='data_input'

SELECT=['','','','']

S_Z=[]
C_Z=[]
G_Z=[]
G_Z=[]
Z_Z=[]

S=[]
C=[]
G=[]
GX=[]
Z=[]
ZZ=[S,C,G,GX,Z]
sort=["","","","",""]
company=[]
service=[]
count=0

result=[]


def get_file_path(path):
    file_paths = []
    for filename in os.listdir(path):
        file_path = os.path.join(path, filename)
        if os.path.isfile(file_path):
            file_paths.append(file_path)
    return file_paths

def open_file(path):
    df = pd.read_excel(path,header=0)
    return df

def judge_zhiju(df):
    S.clear()
    C.clear()
    GX.clear()
    G.clear()
    Z.clear()
    for index, element in df.iterrows():
        if element[SELECT[1]] == '':
            continue
        if element[SELECT[2]] in S_Z:
                S.append(element)
        elif element[SELECT[2]] in C_Z:
                C.append(element)
        elif element[SELECT[2]] in G_Z:
                GX.append(element)
        elif element[SELECT[2]] in G_Z:
                G.append(element)
        elif element[SELECT[2]] in Z_Z:
                Z.append(element)
    pd.DataFrame(S)
    pd.DataFrame(C)
    pd.DataFrame(GX)
    pd.DataFrame(G)
    pd.DataFrame(Z)
    return S,C,GX,G,Z

def compute(ZJS):
    count = 0
    result.clear()
    for ZJ in ZJS:
        total = 0
        for row in ZJ:
            total = total + row[SELECT[3]]
        result.append([sort[count],int(total)])
        count = count + 1
    return result

def save(ZJS,path):
    count=0

    with pd.ExcelWriter('result/'+os.path.basename(path),mode='w') as writer:
            df=pd.DataFrame(compute(ZJS),columns=(['','']))
            df.to_excel(writer, index=False, sheet_name='')
            for ZJ in ZJS:
                df2=pd.DataFrame(ZJ)
                df2.to_excel(writer, index=False,sheet_name=sort[count])
                count=count+1



if __name__ == '__main__':
    paths=get_file_path(File)

    for element in paths:
        judge_zhiju(open_file(element))
        if os.path.exists(os.path.basename(element)):
            os.remove(os.path.basename(element))
        save(ZZ,element)






#print(df)