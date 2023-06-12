import pandas as pd
import numpy as np
from sklearn import preprocessing, linear_model
from sklearn.model_selection import train_test_split
import matplotlib.pyplot as plt
import os, openpyxl

Alpha = 0.4
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# load sample.xlsx
book = openpyxl.load_workbook("./data/sample.xlsx")
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
y = mm.fit_transform(y)
X_reg = np.insert(X_reg,0,1,axis=1)

# make DataFrame
df = pd.DataFrame(data= np.concatenate([X_reg,y],axis=1) ,columns=[0,1,2,3,4,5,6,7,8,9,10,11,"y"])
print(df.head(5))
print(y)

# make Linear regresssion model
X_train, X_test, y_train, y_test = train_test_split(X_reg, y, random_state=0)
print(X_train.shape, y_train.shape)
model = linear_model.Ridge(alpha=Alpha)

model.fit(X_train, y_train)

# show the results
y_pred = model.predict(X_test)
day = [int(i+1) for i in range(y.shape[0])]
x = day
plt.plot(x,y,color="blue")
plt.plot(x[:y_train.shape[0]],y_train[:,0],color="green",label="Training_set")
plt.plot(x[y_train.shape[0]:], y_pred[:,0] ,color="red",label="Prediction")
plt.legend(("Sample_data","Training_set","Prediction"),loc="center left")
plt.show()
