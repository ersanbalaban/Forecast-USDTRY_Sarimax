import warnings
import matplotlib.pyplot as plt
import xlsxwriter
import pandas as pd
import statsmodels.api as sm
import matplotlib
import os
from tkinter import *
from tkinter.filedialog import askopenfilename

# Default Settings and Values
warnings.filterwarnings("ignore")
plt.style.use('fivethirtyeight')
matplotlib.rcParams['axes.labelsize'] = 14
matplotlib.rcParams['xtick.labelsize'] = 12
matplotlib.rcParams['ytick.labelsize'] = 12
matplotlib.rcParams['text.color'] = 'k'
y = ""
forecast = ""
fc_confidence = ""
path = ""
usefilename = ""

# Forecast Process
def chooseFile():
    global y, forecast, fc_confidence, path, usefilename
    path = askopenfilename()
    df = pd.read_excel(path, parse_dates=["Tarih"], index_col="Tarih")
    y = df['Kur'].resample('MS').mean()
    model = sm.tsa.statespace.SARIMAX(y,
                                      order=(1, 1, 1),
                                      seasonal_order=(1, 1, 0, 12),
                                      enforce_stationarity=False,
                                      enforce_invertibility=False)
    results = model.fit()
    forecast = results.get_forecast(steps=100)
    fc_confidence = forecast.conf_int()
    usefilename += path.rsplit('/', 1)[-1]
    labelSelectPath.configure(text=path)


# Draw the graph of the forecast values
def makeGraph():
    ax = y.plot(label='Veriler', figsize=(14, 7))
    forecast.predicted_mean.plot(ax=ax, label='Tahmin')
    ax.fill_between(fc_confidence.index,
                    fc_confidence.iloc[:, 0],
                    fc_confidence.iloc[:, 1], color='k', alpha=.25)
    ax.set_xlabel('Tarih')
    ax.set_ylabel('Tahmin Değeri')
    plt.legend()
    filename = "Forecast_Graph_"
    filename += usefilename.rsplit('.', 1)[0]
    filename += ".png"
    plt.savefig(filename)
    labelMakeGraph.configure(text="Grafik kaydedildi: " + filename)
    plt.show()


# Create an xlsx data of the forecast values
def makeData():
    filename = "Forecast_Data_ " + usefilename
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Tarih")
    worksheet.write(0, 1, "Tahmini Değer")
    worksheet.write(0, 2, "En Düşük Tahmin")
    worksheet.write(0, 3, "En Yüksek Tahmin")
    row = 1
    col = 0
    for x in forecast.predicted_mean.items():
        dt_obj = x[0].to_pydatetime().date().strftime('%m-%Y')
        worksheet.write(row, col, str(dt_obj))
        worksheet.write(row, col + 1, str(round((x[1]), 2)))
        worksheet.write(row, col + 2, str(round(fc_confidence.iloc[:, 0][1], 2)))
        worksheet.write(row, col + 3, str(round(fc_confidence.iloc[:, 1][1], 2)))
        row += 1
    row = 1
    for y in fc_confidence.iloc[:, 0].items():
        worksheet.write(row, col + 2, str(round((y[1]), 2)))
        row += 1
    row = 1
    for z in fc_confidence.iloc[:, 1].items():
        worksheet.write(row, col + 3, str(round((z[1]), 2)))
        row += 1
    workbook.close()
    labelMakeData.configure(text="Tablo kaydedildi: " + filename)
    os.startfile(filename)


# Graphical UI
window = Tk()
window.geometry("800x400")
window.title("DPU || Döviz Kuru Tahmini")
buttonSelectPath = Button(window, text="Xlsx Dosyası Seç", command=lambda: chooseFile())
buttonSelectPath.place(x=150, y=50)
labelSelectPath = Label(window, text="Önce dosya seçiniz")
labelSelectPath.place(x=300, y=50)
buttonMakeGraph = Button(window, text="Tahmin Grafiği Oluştur", command=lambda: makeGraph())
buttonMakeGraph.place(x=115, y=150)
labelMakeGraph = Label(window, text="Grafik oluşturmak için butona tıklayınız.")
labelMakeGraph.place(x=300, y=150)
buttonMakeData = Button(window, text="Tahmin Verilerinden Tablo Oluştur", command=lambda: makeData())
buttonMakeData.place(x=55, y=250)
labelMakeData = Label(window, text="Tablo (xlsx) oluşturmak için butona tıklayınız.")
labelMakeData.place(x=300, y=250)
window.mainloop()
