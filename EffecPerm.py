from contextlib import redirect_stderr
from pickle import TRUE
import pandas as pd 
import numpy as np
from sklearn import linear_model
from sklearn import metrics
import statsmodels.api as sm
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import seaborn as sn
import tkinter as tk 
#import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
#from IPython.display import display_html
import math
s=r'‪C:\Users\aayushi\Desktop\input.xlsx'
s = s.lstrip('\u202a')
xlsx = pd.ExcelFile(s)
df1=pd.read_excel(xlsx,'Sheet3')
df10=pd.read_excel(xlsx,'Sheet2')
col_mapping = [f"{c[0]}:{c[1]}" for c in enumerate(df1.columns)]
cols = ['Perm']  # We don't want to convert the Final grade column.
for col in cols:  # Iterate over chosen columns
  df1[col] = [float(str(val).replace('. ','.')) for val in df1[col].values]
def EffectPerm(x):
    z=1118*(2*0.1454*(x-118)+x+2*118)/(2*118+x-0.1454*(x-118))
    return round((z),1)
df1.insert(3,'EffecP',np.nan)
for i in range(0,14):
  df2=df1.iloc[i,1].astype(float)

  
 
  

  df4 = EffectPerm(df2)
  df1.iloc[i,3]=df4

#print(df1)

def apc1(x1):
    z1=(6.18)*(np.cbrt((2)/(x1-1)))
    #z11=1/x1
    #z1=(6.18)*(np.cbrt(1+3*z11+3*z11**2+2*z11**3))
    
    return round((z1),3)
df1.insert(4,'apc1',np.nan)
for i in range(0,14):
  df3=df1.iloc[i,1].astype(float)

  
 
  

  df5 = apc1(df3)
  df1.iloc[i,4]=df5

df1.insert(7,'V(n)',np.nan)

def EffectV(x,y):
    if(x==2 and y==4):
        return 48
    elif(x==1 and y==5):
        return 30
    elif(x==3 and y==3):
        return 54

for i in range(0,14):
  df6=df1.iloc[i,5].astype(float)
  df7=df1.iloc[i,6].astype(float)

  
 
  

  df8 = EffectV(df6,df7)
  df1.iloc[i,7]=df8
def ra(x2):
   for i in range(0,78):
      E=df10.iloc[i,4]
      if(x2==E):
        return (df10.iloc[i,0].astype(float))


def nval(x2):
    for i in range(0,78):
      E=df10.iloc[i,4]
      if(x2==E):
        return (df10.iloc[i,2].astype(float))
def rb(y2):
   for i in range(0,78):
      E=df10.iloc[i,4]
      if(y2==E):
        return (df10.iloc[i,0].astype(float))

      
df1.insert(10,'rav',np.nan)
for i in range(0,14):
  df9=df1.iloc[i,8]
  df11=df1.iloc[i,8]
  
  rax=ra(df9)
  ray=rb(df11)
  raz=df1.iloc[i,4]
  

  df1.iloc[i,10]=round(float((rax+ray+raz)/3),3)
def alat(x22,y22):
       for i in range(0,14):
         z22=2.45*(x22)**(0.09)*y22
         return round(z22,3)

      
df1.insert(11,'aeff',np.nan)
for i in range(0,14):
  df12=df1.iloc[i,7]
  df13=df1.iloc[i,10]
  aeff=alat(df12,df13)
  
  

  df1.iloc[i,11]=aeff

def ueff(x2,y2):
       for i in range(0,14):
         ud=1.6/x2
         w2=(1.53*10**10)**2
         c2=(3*10**8)**2
         wc2=w2/c2
         cmm=0.1454/10
         ueff=ud*(1+cmm*wc2*x2*(y2**2))
         return round(ueff,3)

df1.insert(12,'ueff',np.nan)
for i in range(0,14):
  df14=df1.iloc[i,1]
  df15=df1.iloc[i,11]
  uefff=ueff(df14,df15)
  
  

  df1.iloc[i,12]=uefff

def zin(x2,y2):
       for i in range(0,14):
         zin=np.sqrt((y2*4*3.14*10**(-7))/(x2*8.85*10**(-12)))
         return round(zin,3)

df1.insert(13,'zin',np.nan)
for i in range(0,14):
  df16=df1.iloc[i,3]
  df17=df1.iloc[i,12]
  zinn=zin(df16,df17)
  
  

  df1.iloc[i,13]=zinn

def Rl(x2):
       for i in range(0,14):
         rl1=20*np.log10(np.abs((x2-460)/(x2+460)))
         return round(rl1,3)

df1.insert(14,'RL',np.nan)
for i in range(0,14):
  df18=df1.iloc[i,13]
  rl2=Rl(df18)
  
  

  df1.iloc[i,14]=rl2

my_list={'Material':['CaTiO3',
         'CaZrO3',
         'SrZrO3',
         'BaZrO3',
         'LaGaO3',
         'SrTiO3',
         'NdAlO3',
         'LaAlO3',
         'PrAlO3',
         'ErAlO3',
         'DyAlO3',
         'GdAlO3',
         'SmAlO3',
         'YAlO3']}
df5=pd.DataFrame(my_list)
#n=df1.columns[0]
#df1.drop(n, axis = 1, inplace = True)
#df1[n]=my_list
df1.iloc[:,0]=df5
#print(df1)
p=r'‪C:\Users\aayushi\Desktop\output2d.xlsx'
p = p.lstrip('\u202a')
writer = pd.ExcelWriter(p, engine = 'xlsxwriter')
df1.to_excel(writer, sheet_name = 'SheetC')

#writer.close()


df5.loc[:,'ra']=df1.loc[:,'ra']
df5.insert(2,'ra_val',np.nan)
for i in range(0,14):
      df99=df5.loc[i,'ra']
     
  
      raxx=ra(df99)
      naxx=nval(df99)
      df5.iloc[i,2]=raxx
      #df5.iloc[i,2]=raxx
  

df5.loc[:,'rb']=df1.loc[:,'rb']

df5.insert(4,'rb_val',np.nan)
for i in range(0,14):
      df101=df5.loc[i,'rb']
     
  
      rayy=ra(df101)
      nayy=nval(df101)
      df5.iloc[i,4]=rayy
df5.loc[:,'RL']=df1.loc[:,'RL']
#print(df5)
#p=r'‪C:\Users\aayushi\Desktop\output2d.xlsx'
#p = p.lstrip('\u202a')
#writer = pd.ExcelWriter(p, engine = 'xlsxwriter')
#df5.to_excel(writer, sheet_name = 'SheetD')

#import sys
 
#file_path = r'C:\Users\aayushi\Desktop\Results.txt'
#file_path=file_path.lstrip('\u202a')
#sys.stdout = open(file_path, "w")
#print("This text will be added to the file")

xi=pd.DataFrame(df5.loc[:,['ra_val','rb_val']])
#print(xi)
yi=pd.DataFrame(df5.loc[:,'RL'])
from sklearn.model_selection import train_test_split
#xis = sm.add_constant(xi)
#x_train, x_test, y_train, y_test = train_test_split(xi, yi, test_size = 0.6, random_state = 60)
x_train, x_test, y_train, y_test = train_test_split(xi, yi)
 # adding a constant
 
#regr = sm.OLS(y_train, x_train).fit()
regr = linear_model.LinearRegression()
regr.fit(xi, yi)

print('Intercept: \n', regr.intercept_)
print('Coefficients: \n', regr.coef_)
y_pred_mlr= regr.predict(x_test)
y_pred_all=regr.predict(xi)
df5['RL2']=y_pred_all
print(df5)
print(df1)

#i1=df1.loc[df1['ra']=='Ca']
for i in range(0,14):
 
 P1=df1.iloc[i,3]
#i2=df1.loc[df1['rb']=='Al']
 P2=df1.loc[4,'Perm']
 Pav=(P1+P2)/2
 Peff=EffectPerm(Pav)
 print(round(Peff),2)
#Predicted values
#print("Prediction for test set: {}".format(y_pred_mlr))
#mlr_diff = pd.DataFrame({'Actual value': [y_test], 'Predicted value': [y_pred_mlr]})

#y_test['RL2']=y_pred_mlr.tolist()
#dff=pd.DataFrame(y_test)
#dff['RL2']=y_pred_mlr
#print(dff)
#dff=pd.concat([y_test,y_pred_mlr],axis=1)
#dff.rename(columns={'RL':'RL_ACTUAL',0:'RL_PREDICTED'},inplace=TRUE)
#print(dff)
#y_test_styler = y_test.style.set_table_attributes("style='display:inline'").set_caption('df1')
#y_pred_mlr_styler = y_pred_mlr.style.set_table_attributes("style='display:inline'").set_caption('df2')
#df2_t_styler = df2.T.style.set_table_attributes("style='display:inline'").set_caption('df2_t')

#display_html(y_test_styler._repr_html_()+y_pred_mlr_styler._repr_html_(), raw=True)
#print("%$ %$ " %{y_test,y_pred_mlr})
#print(y_pred_mlr)

meanAbErr = metrics.mean_absolute_error(y_test, y_pred_mlr)
meanSqErr = metrics.mean_squared_error(y_test, y_pred_mlr)
rootMeanSqErr = np.sqrt(metrics.mean_squared_error(y_test, y_pred_mlr))
#file_path = r'C:\Users\aayushi\Desktop\Results.txt'
#file_path=file_path.lstrip('\u202a')
#sys.stdout = open(file_path, "w")
#print_model = regr.summary()
#print(print_model)
#print(regr.pvalues)
print('R squared: {:.2f}'.format(regr.score(xi,yi)*100))
print('Mean Absolute Error:', meanAbErr)
print('Mean Square Error:', meanSqErr)
print('Root Mean Square Error:', rootMeanSqErr)

#y_test['RL2']=y_pred_mlr.tolist()
dff=pd.DataFrame(y_test)
dff['RL2']=y_pred_mlr
print(dff)
#plt.scatter(df5['ra'], df5['RL'], color='red')
#data=sn.load_dataset(df5)
ax=sn.regplot(x='ra_val', y='RL',data=df5,color='red')
plt.title('Variation of Reflection Loss vs cation A radius', fontsize=14)
plt.xlabel('Radius of cation A (Å) ', fontsize=14)
plt.ylabel('Reflection Loss(dB)', fontsize=14)
plt.show()

ax1=sn.regplot(x='rb_val', y='RL',data=df5,color='green')
plt.title('Variation of Reflection Loss vs cation B radius', fontsize=14)
plt.xlabel('Radius of cation B (Å) ', fontsize=14)
plt.ylabel('Reflection Loss(dB)', fontsize=14)
plt.show()
#plt.plot(df5['ra'], df5['RL'], color='green')
#for i, txt in enumerate(df5['ra']):
#    plt.annotate(txt, (xl[i], df5.loc[i,'RL']))
#plt.grid(True)
#plt.show()

X = df5[['ra_val', 'rb_val']].values.reshape(-1,2)
Y = df5['RL2']

######################## Prepare model data point for visualization ###############################

x = X[:, 0]
y = X[:, 1]
z = Y
xx_pred = np.linspace(1, 5, 40)  # range of price values
yy_pred = np.linspace(1, 5, 40)  # range of advertising values
xx_pred, yy_pred = np.meshgrid(xx_pred, yy_pred)
model_viz = np.array([xx_pred.flatten(), yy_pred.flatten()]).T

# Predict using model built on previous step
ols = linear_model.LinearRegression()

model = ols.fit(X, Y)
predicted = model.predict(model_viz)
#r2 = regr.score(X, Y)
plt.style.use('default')

fig = plt.figure(figsize=(12, 4))

ax1 = fig.add_subplot(131, projection='3d')
ax2 = fig.add_subplot(132, projection='3d')
ax3 = fig.add_subplot(133, projection='3d')

axes = [ax1, ax2, ax3]

for ax in axes:
    ax.plot(x, y, z, color='k', zorder=15, linestyle='none', marker='o', alpha=0.5)
    ax.scatter(xx_pred.flatten(), yy_pred.flatten(), predicted, facecolor=(0,0,0,0), s=20, edgecolor='#70b3f0')
    ax.set_xlabel('Cation A Radius (Å)', fontsize=12)
    ax.set_ylabel('Cation B Radius(Å)', fontsize=12)
    ax.set_zlabel('Reflection Loss(dB)', fontsize=12)
    ax.locator_params(nbins=4, axis='x')
    ax.locator_params(nbins=5, axis='x')

ax1.view_init(elev=25, azim=-60)
ax2.view_init(elev=15, azim=15)
ax3.view_init(elev=25, azim=60)

fig.suptitle('Reflection Loss vs Atomic Radii Model Visualization ' , fontsize=15, color='k')

fig.tight_layout()
plt.show()
#p=r'C:\Users\aayushi\Desktop\outputper
# m.xlsx'
#p = p.lstrip('\u202a')

#writer = pd.ExcelWriter(p, engine = 'xlsxwriter')
#df5.to_excel(writer, sheet_name = 'SheetA')
#p=r'‪C:\Users\aayushi\Desktop\output.xlsx'
#p = p.lstrip('\u202a')
#writer = pd.ExcelWriter(s, engine = 'xlsxwriter')
#df1.to_excel(writer)
# tkinter GUI
root= tk.Tk()

canvas1 = tk.Canvas(root, width = 500, height = 300)
canvas1.pack()

# with sklearn
Intercept_result = ('Intercept: ', regr.intercept_)
label_Intercept = tk.Label(root, text=Intercept_result, justify = 'center')
canvas1.create_window(260, 220, window=label_Intercept)

# with sklearn
Coefficients_result  = ('Coefficients: ', regr.coef_)
label_Coefficients = tk.Label(root, text=Coefficients_result, justify = 'center')
canvas1.create_window(260, 240, window=label_Coefficients)

# New_Interest_Rate label and input box
label1 = tk.Label(root, text='Type Cation A radius (1-5 Å) ')
canvas1.create_window(100, 100, window=label1)

entry1 = tk.Entry (root) # create 1st entry box
canvas1.create_window(270, 100, window=entry1)

# New_Unemployment_Rate label and input box
label2 = tk.Label(root, text=' Type Cation B radius (Å) ')
canvas1.create_window(120, 120, window=label2)

entry2 = tk.Entry (root) # create 2nd entry box
canvas1.create_window(270, 120, window=entry2)

def values(): 
    global ra_vale #our 1st input variable
    ra_vale = float(entry1.get()) 
    
    global rb_vale #our 2nd input variable
    rb_vale = float(entry2.get()) 
    
    Prediction_result  = ('Predicted Stock Index Price: ', regr.predict([[ra_vale ,rb_vale]]))
    label_Prediction = tk.Label(root, text= Prediction_result, bg='orange')
    canvas1.create_window(260, 280, window=label_Prediction)

def compound(): 
    global xa #our 1st input variable
    xa = float(entry1.get()) 
    
    global ya #our 2nd input variable
    ya = float(entry2.get()) 

    xan = xa_name(xa)
    xbn = xb_name(ya)
    perm=Permtt(xan,xbn)
    
    Compound_result  = ('Predicted Compound Name ',str(xan),str(xbn),'O3')
    Compound_Perm  = ('Predicted Compound Effective Pemittivity ',perm)
    label_compound = tk.Label(root, text= Compound_result, bg='yellow')
    label_compound_perm = tk.Label(root, text= Compound_Perm, bg='yellow')
    #canvas1.create_window(280, 300, window=label_compound)
    canvas1.create_window(260, 280, window=label_compound)
    canvas1.create_window(290, 300, window=label_compound_perm)
def Permtt(xa11,xb11):
  for i in range(0,14):
    
      dfa=df1.iloc[i,8]
      dfb=df1.iloc[i,9]
    
      if(xa11==dfa)and(xb11==dfb):
       return(round(df1.iloc[i,3],2))
    
  for j in range(0,14):
    dfa1=df1.iloc[j,8]
    if(dfa1==xa11):
      P1=df1.iloc[j,3]
  for j in range(0,14):
    dfb1=df1.iloc[j,9]
    if(dfb1==xb11):
      P2=df1.iloc[j,3]
     
  Pav=(P1+P2)/2
  Peff=EffectPerm(Pav)
  return(Peff)


def xa_name(xa1):
  if(1<=xa1<=1.3):
    return ('Er')
  elif(1.3<xa1<=1.5):
    return ('Dy')
  elif(1.5<xa1<2):
      return ('Gd')
  elif(2<=xa1<=2.5):
      return ('Nd')
  elif(2.5<xa1<=2.8):
      return ('Ca')
  elif(2.8<=xa1<=3.5):
      return('Y')
  elif(3.5<xa1<=3.6):
      return ('Sr')
  elif(3.6<xa1<=4.2):
      return ('La')
  elif(4.2<xa1<=5):
      return ('Ba')

def xb_name(xb1):
  if(1<=xb1<=1.4):
    return ('Al')
  elif(1.4<xb1<=2.2):
    return ('Ga')
  elif(2.2<xb1<3):
      return ('Ti')
  elif(3<=xb1<=4):
      return ('Zr')
 



            

 

    
button1 = tk.Button (root, text='Predict Reflection Loss',command=values, bg='orange')
 # button to call the 'values' command above 
button2 = tk.Button (root, text='Predict Compound Name',command=compound, bg='orange')
canvas1.create_window(270, 150, window=button1)
canvas1.create_window(285, 185, window=button2)
 
#plot 1st scatter 
figure3 = plt.Figure(figsize=(5,4), dpi=100)
#figure3 = plt.subplots(ncols=1, sharey=True)
ax3 = figure3.add_subplot(111)
sn.regplot(x=df5['ra_val'], y=df5['RL'], ax=ax3)
#sns.regplot(x=idx, y=df['y'], ax=ax2)

#ax=sn.regplot(x='ra_val', y='RL',data=df5,color='red')
#plt.title('Variation of Reflection Loss vs cation A radius', fontsize=14)
#plt.xlabel('Radius of cation A (Å) ', fontsize=14)
#plt.ylabel('Reflection Loss(dB)', fontsize=14)
#plt.show()
#ax3 = figure3.add_subplot(111)
#ax3.regplot(df5['ra_val'], df5['RL'],color='red')
#ax3.scatter(df5['ra_val'].astype(float),df5['RL'].astype(float), color = 'r')
scatter3 = FigureCanvasTkAgg(figure3, root) 
scatter3.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH)
#plt.title('Variation of Reflection Loss vs cation A radius', fontsize=14)
#plt.xlabel('Radius of cation A (Å) ', fontsize=14)
#plt.ylabel('Reflection Loss(dB)', fontsize=14)
#plt.show()
ax3.legend(['Reflection Loss(dB)']) 
ax3.set_xlabel('Radius of cation A (Å)')
ax3.set_title('Variation of Reflection Loss vs cation A radius')

#plot 2nd scatter 
figure4 = plt.Figure(figsize=(5,4), dpi=100)
ax4 = figure4.add_subplot(111)
ax4.scatter(df5['rb_val'].astype(float),df5['RL'].astype(float), color = 'g')
scatter4 = FigureCanvasTkAgg(figure4, root) 
scatter4.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH)
ax4.legend(['Stock_Index_Price']) 
ax4.set_xlabel('Unemployment_Rate')
ax4.set_title('Unemployment_Rate Vs. Stock Index Price')

root.mainloop()
writer.save()

