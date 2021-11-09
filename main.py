import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
excel_url = './Office Table.xlsx'

result=[]
for j in range(4):
  
  df = pd.read_excel(excel_url,usecols=[2],sheet_name=j)

  E=[]
  for i in range(len(df.index)):  # E값 가져오기
    E.append(df["E.1"][i])

  df2 = pd.read_excel(excel_url,usecols = "D:BF") #자료값 가져오기
  my_list=df2.to_numpy()              # 2차원 리스트로 변환
  pinv_my_list=np.linalg.pinv(my_list) # Pseudo inverse

  c=pinv_my_list.dot(E)               # c값구하기 pinv(A)*Ac=pinv(A)*E
  #print(len(c))
  result.append(c)

result_excel=pd.DataFrame(result)      # c값데이터 타입변환 
result_excel.to_excel('result.xlsx') # c값 엑셀로 뽑기

X = np.linspace(0, 163, 163)
Ep=my_list.dot(c)

result_excel_E=pd.DataFrame(Ep)
result_excel_E.to_excel('calculated_E.xlsx')
plt.scatter(Ep, E,color='#EDAA7D')

error = (E-Ep)/E*100
print(round(abs(error.sum()/len(error)),2),'%')
plt.show()
