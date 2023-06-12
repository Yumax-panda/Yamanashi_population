
import pandas as pd
import numpy as np
from sklearn import preprocessing
from sklearn import linear_model
import matplotlib.pyplot as plt
import os,openpyxl

Alpha = 4

cwd = os.getcwd()
# load sample.xlsx
book = openpyxl.load_workbook(cwd+"\\data\\sample.xlsx")
sheet = book["Sheet"]
X = []
for i in range(132):
    for j in range(11):
        X.append(float(sheet.cell(row=i+1, column=j+1).value))
X = np.array(X).reshape(-1,11)

# load population.xlsx
book = openpyxl.load_workbook('population.xlsx')
sheet = book["Sheet"]
y = []
for i in range(132):
    y.append(int(sheet.cell(row=i+1, column=1).value))
y = np.array(y).reshape(132,1)

# regularization
mm = preprocessing.MinMaxScaler()
X_reg = mm.fit_transform(X)
X_reg = np.insert(X_reg,0,1,axis=1)
#merge data [X,y]
Data = np.concatenate([X_reg,y],axis=1)

# make DataFrame
df = pd.DataFrame(data= Data,columns=[0,1,2,3,4,5,6,7,8,9,10,11,"y"])
print(df.head(5))

# make Linear regresssion model
X_train, X_test, Y_train, Y_test = df.iloc[:99,:11],df.iloc[100:,:11],df.iloc[:99,-1],df.iloc[100:,-1]
model = linear_model.Ridge(alpha=Alpha)

model.fit(X_train, Y_train)

# show the results
y_pred = model.predict(X_reg[:,:11])
y_val = model.predict(X_test)
day = [int(i+1) for i in range(132)]
x = day
plt.plot(x,y,color="blue")
plt.plot(x,y_pred,color="green",label="Training_set")
plt.plot(x[100:],y_val,color="red",label="Prediction")
plt.legend(("Sample_data","Training_set","Prediction"),loc="center left")

plt.savefig(f"output{Alpha}.pdf")
plt.show()