import openpyxl
import matplotlib.pyplot
import numpy

day_num=[]
workbook1=openpyxl.load_workbook('covid_19_confirmed_time_series_used.xlsx')
worksheet1 = workbook1.active
for cell in worksheet1['A']:
    day_num.append(cell.value)
print (day_num)

confirmed_num=[]
for cell in worksheet1['B']:
    confirmed_num.append(cell.value)
print (confirmed_num)

x = day_num
y = confirmed_num


params = numpy.polyfit(x, y, 3)
print(params)
param_func = numpy.poly1d(params)
print(param_func)

y_predict = param_func(x)

matplotlib.pyplot.xlabel('时间（第n天）', fontproperties = 'SimSun', fontsize = 13, color = 'black')
matplotlib.pyplot.ylabel('确诊人数(千万人)', fontproperties = 'SimSun', fontsize = 13, color = 'black')
matplotlib.pyplot.title(r'Covid-19确诊人数预测', fontproperties = 'SimHei', fontsize = 20)

matplotlib.pyplot.plot(x, y, color='m', linestyle='None',marker='o',label=u'data')
matplotlib.pyplot.plot(x, y_predict, color='b',label=u'Fit curving')
matplotlib.pyplot.legend(['data','Fit curving'])
matplotlib.pyplot.show()

y_predict_array = numpy.array(y_predict)
y_array = numpy.array(y)
MSE=numpy.sum(numpy.power((y_predict_array.reshape(-1, 1) - y_array), 2))/len(y)
print("MSE:", MSE)
