from flask import Flask, render_template, request, send_from_directory, redirect, url_for
from werkzeug.utils import secure_filename
from math import *
import os
import pandas as pd
import plotly.express as px
import plotly.io as pio
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import shutil

app = Flask(__name__, static_folder='static')
os.chdir(os.path.dirname(__file__))
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['DOWNLOAD_FOLDER'] = 'downloads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsm','xlsx'}

VRP_MO=[4206.506,4290.2613,4644.635,5196.1364,5406.0765,6809.9511,7720.8426]#2016 2017 2018 2019 2020 2021 2022

def sankey(table):
  year=table.iloc[0,10]
  table = table.iloc[2:].reset_index(drop=True)
  names = table.iloc[:,0].head(20).tail(18).tolist()+['Баланс']
  b= len(names) - 1
  v = table.iloc[:,11].head(20).tail(18).astype('str').str.replace('-','').tolist()
  v[6]=str(round(table.iloc[8,9],2))
  v[7]=str(round(table.iloc[9,10],2))
  names[0]='Производство'  
  names[1]='Импорт'
  names[4]='Потр. первичной энергии'
  names[5]='Стат. расхождение'
  names[7]='Производство ТЭ'
  names[10]='Потери'
  names[11]='Преобр. топлива'
  names[17]='Конечное потр.'
  names[-1]+=' ' + str(round((float(v[0])+float(v[1]))/1000,2)) 
  for i in range(len(names)-1):
      names[i]+=' '+str(round(float(v[i])/1000,2))  
  fig = go.Figure(data=[go.Sankey(
      arrangement = "snap",
      node = dict(
          pad = 15,
          thickness = 10,
          line = dict(color = "black", width = 1),
          label = names,      
          x = [0.3, 0.3, 0.5, 0.5, 0.5, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.4], # 0 пэр 1 ввоз 2 вывоз 3 изм зап 4 ппэ 5 стат расх 6 произв ээ 7 произв тэ 8 преобр топ 9 соб нуж 10 потери 11 кп 12 бал
          y = [0.8, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1],
      ),
      link = dict(
          source = [0, 1, b, b, b, 4, 4, 4,4,4,4,4],
          target = [b, b, 2, 3, 4, 5, 6, 7,11,15,16,17],
          value = [v[0],v[1],v[2],v[3],v[4],v[5],v[6],v[7],v[11],v[15],v[16],v[17]]
  ))])
  fig.update_layout(
        title={
            'text': "Структура первичного потребления ТЭР за "+year,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=12
        )
    )
  return fig

def sankey2(table):
  names =['Производство', 'Импорт','Доступно из всех ист.','Прямой перенос','Трансф.','Транф. убытки','Доступно после трансф.','Транзит','Конечное потр.']
  year=table.iloc[0,10]
  table = table.iloc[2:].reset_index(drop=True)
  v = table[11].head(20).tail(18).astype('str').str.replace('-','').tolist()
  v = [eval(i) for i in v]
  vDop= [int(table[4][19]),int(table[5][19]),int(table[9][19]),int(table[10][19])]
  val = [v[0], v[1], vDop[1]+v[2], v[6]+v[7]+v[11]+v[15]+v[16] + vDop[0]+vDop[2]+vDop[3], vDop[1]+v[2], v[6]+v[7]+v[11]+v[15]+v[16], vDop[0]+vDop[2]+vDop[3], v[2], v[17]] 
  names[0]+=' '+ str(round(float(val[0])/1000,2))  
  names[1]+=' '+ str(round(float(val[1])/1000,2)) 
  names[2]+=' '+ str(round((float(val[0])+float(val[1]))/1000,2)) 
  names[3]+=' '+ str(round(float(val[4])/1000,2)) 
  names[4]+=' '+ str(round(float(val[3])/1000,2)) 
  names[5]+=' '+ str(round(float(val[5])/1000,2)) 
  names[6]+=' '+ str(round((float(val[3])+float(val[4])-float(val[5]))/1000,2)) 
  names[7]+=' '+ str(round(float(val[7])/1000,2)) 
  names[8]+=' '+ str(round(float(val[8])/1000,2)) 
  fig = go.Figure(data=[go.Sankey(
    arrangement = "snap",
    node = dict(
      pad = 15,
      thickness = 10,
      line = dict(color = "black", width = 1),
      label = names,
      x = [0.2, 0.2, 0.3, 0.5, 0.5, 0.65, 0.65, 0.85, 0.85],
      y = [0.1, 0.1, 0.1, 0.1, 0.4, 0.8, 0.1, 0.1, 0.45],
    ),
    link = dict(
      source = [0, 1, 2, 2, 3, 4, 4, 6, 6],
      target = [2, 2, 3, 4, 6, 5, 6, 7, 8],
      value =  val
  ))])
  fig.update_layout(
        title={
            'text': "Пропорции ТЭБ "+year,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=12
        )
    )
  return fig

def piePPE(table): # потребление первичной энергии
  year=table.iloc[0,10]
  table = table.iloc[2:].reset_index(drop=True)
  labels = table.iloc[0].to_list()[2:11]
  values = table.iloc[6].values.tolist()[2:11]  
  fig = go.Figure(data=[go.Pie(labels=labels,values=values)])
  fig.update_layout(
    title={
      'text': "Потребление первичной энергии за "+year,
      'y':0.9,
      'x':0.5,
      'xanchor': 'center',
      'yanchor': 'top'
    },
    font=dict(
      family="Arial, sans-serif",
      size=18
    ))
  return fig

def pieKP(table): # конечное потребление
  year=table.iloc[0,10]
  table = table.iloc[2:].reset_index(drop=True)
  labels = table.iloc[0].to_list()[2:11]
  values = table.iloc[19].values.tolist()[2:11]  
  fig = go.Figure(data=[go.Pie(labels=labels,values=values)])
  fig.update_layout(
        title={
            'text': "Конечное потребление за "+year,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=18
        )
    )
  return fig

def barChart(tables):
  dataBal=[]
  dataPPE=[]
  dataKP=[]
  for table in tables:        
    year=table.iloc[0,10]
    table = table.iloc[2:].reset_index(drop=True)
    dataBal.append([year[:4], int(table.iloc[:,11][2])+int(table.iloc[:,11][3])])
    dataPPE.append([year[:4], int(table.iloc[:,11][6])])
    dataKP.append([year[:4], int(table.iloc[:,11][19])])
  dfBal=pd.DataFrame(dataBal,columns=['Год','Баланс'])
  dfPPE=pd.DataFrame(dataPPE,columns=['Год','Потребление первичной энергии'])
  dfKP=pd.DataFrame(dataKP,columns=['Год','Конечное потребление'])
  figBal =[
     px.bar(dfBal,x='Год',y='Баланс').update_layout(
        title={
            'text': 'Сравнение балансов',
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=18
        )
    ),
    px.bar(dfPPE,x='Год',y='Потребление первичной энергии').update_layout(
        title={
            'text': 'Сравнение потребления первичной энергии',
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=18
        )
    ),
    px.bar(dfKP,x='Год',y='Конечное потребление').update_layout(
        title={
            'text': 'Сравнение конечного потребления',
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        font=dict(
            family="Arial, sans-serif",
            size=18
        )
    )]
  return figBal

def parcoordsChart(tables, tables2):
  energy_sources = ["Уголь", "Сырая нефть", "Нефтепродукты", "Природный газ", "Твёрдое топливо", "Гидроэнергия", "Атомная энергия", "Электро энергия", "Тепловая энергия"]
  energy_sources2 = ["Уголь", 'Дизель', 'Мазут', 'СУГ', 'Природный газ']
  sectors_sources = ['Сельское хозяйство', 'Промышленность', 'Строительство', 'Транспорт', 'Сфера услуг', 'Население']

  ranges_ppe = [[1, 2500], [0, 3000], [1000, 12000], [20000, 26000], [1, 1100], [100, 300], [0, 300], [3000, 5000], [1000, 1600]]
  ranges_ppe2 = [[500,2000], [1000,4100], [0,500], [0,150], [20000,25000]]
  ranges_kp = [[0, 3000], [0, 3000], [1000, 10000], [4000, 8000], [0, 1000], [0, 1000], [0, 1000], [4000, 5500], [8000, 11000]]
  ranges_ee = [[0, 500], [1500, 2000], [100, 200], [400, 800], [400, 1000],[1000, 1500]]
  # Инициализация данных
  ppe_data = {source: [] for source in energy_sources}
  ppe_data2 = {source: [] for source in energy_sources2}
  kp_data = {source: [] for source in energy_sources}
  ee_data = {source: [] for source in sectors_sources}
  years = []
  colors = [[0, 'red'], [0.5, 'yellow'], [1, 'blue']]

  # Обработка таблиц
  for table in tables:
    year = table.iloc[0, 10]
    years.append(int(year[:4]))
    table = table.iloc[2:].reset_index(drop=True)

    for idx, source in enumerate(energy_sources, start=2):
      ppe_value = table.iloc[6,idx] if pd.notna(table.iloc[6,idx]) else 0
      kp_value = table.iloc[19,idx] if pd.notna(table.iloc[19,idx]) else 0
      
      ppe_data[source].append(int(ppe_value))
      kp_data[source].append(int(kp_value))

    for i in range(len(sectors_sources)):
      if i==0:
        ee_data[sectors_sources[i]].append(int(table.iloc[20][9]))
      elif i==1:
        ee_data[sectors_sources[i]].append(int(table.iloc[21][9]))
      elif i==2:
        ee_data[sectors_sources[i]].append(int(table.iloc[27][9]))
      elif i==3:
        ee_data[sectors_sources[i]].append(int(table.iloc[28][9]))
      elif i==4:
        ee_data[sectors_sources[i]].append(int(table.iloc[33][9]))
      elif i==5:
        ee_data[sectors_sources[i]].append(int(table.iloc[34][9]))

  for table in tables2:
    for i in range(len(energy_sources2)):
      if i==0:
        ppe_data2[energy_sources2[i]].append(int(table['Уголь'].iloc[6,1]))
      elif i==1:
        ppe_data2[energy_sources2[i]].append(int(table['Дизель'].iloc[6,1]))
      elif i==2:
        ppe_data2[energy_sources2[i]].append(int(table['Мазут'].iloc[6,1]))
      elif i==3:
        ppe_data2[energy_sources2[i]].append(int(table['СУГ'].iloc[6,1]))
      elif i==4:
        ppe_data2[energy_sources2[i]].append(int(table['Пр-й газ'].iloc[6,1]))

  # Функция для создания графика
  def create_figure(data, ranges, title):
    dimensions = [dict(tickvals=years, label='Год', values=years)] + [
      dict(range=ranges[idx], 
           label=source, 
           values=values,
           tickformat=".0f")
      for idx, (source, values) in enumerate(data.items())
    ]
    fig = go.Figure(data=go.Parcoords(
      line=dict(color=years, colorscale=colors),
      dimensions=dimensions,
      unselected=dict(line=dict(color='green', opacity=1))
    ))
    fig.update_layout(
      title=dict(text=title, y=0.95, x=0.5, xanchor='center', yanchor='top'),
      font=dict(family="Arial, sans-serif", size=18)
    )
    return fig

  # Создание графиков
  figParPPE = create_figure(ppe_data, ranges_ppe, "Динамика изменения потребления первичной энергии")
  figParKP = create_figure(kp_data, ranges_kp, "Динамика изменения конечного потребления")
  figParEE = create_figure(ee_data,ranges_ee,'Конечное потребление ЭЭ по секторам экономики')
  figParPPE2 = create_figure(ppe_data2,ranges_ppe2, 'Выборочная динамика изменения потребления первичной энергии')

  return [figParPPE, figParPPE2, figParKP, figParEE]

def heatMapsGeneral(tables):
  resources = ['Уголь', 'Нефтепродукты', 'Природный газ', 'Электроэнергия', 'Тепловая энергия']
  years = []
  ppe_data = {resource: [] for resource in resources}
  kp_data = {resource: [] for resource in resources}

  # Обработка данных из таблиц
  for table in tables:
    year = table.iloc[0, 10]
    years.append(int(year[:4]))
    table = table.iloc[3:].reset_index(drop=True)
    # Индексы колонок соответствующих ресурсов
    resource_indices = {
      'Уголь': 2, 
      'Нефтепродукты': 4, 
      'Природный газ': 5, 
      'Электроэнергия': 9, 
      'Тепловая энергия': 10
    }
    for resource, col_idx in resource_indices.items():
      ppe_data[resource].append(round(table.iloc[5, col_idx], 2))
      kp_data[resource].append(round(table.iloc[18, col_idx], 2))

  # Функция для создания тепловой карты
  def create_heatmap(data, title):
    values = [data[resource] for resource in resources]
    fig = go.Figure(data=go.Heatmap(
      z=values,
      x=years,
      y=resources,
      text=values,
      texttemplate="%{text}",
      colorscale='YlGnBu',
      colorbar=dict(title="Потребление")
    ))
    fig.update_layout(
      xaxis_title='Годы',
      yaxis_title='Ресурсы',
      xaxis=dict(tickvals=years, ticktext=[str(year) for year in years]),
      yaxis=dict(tickvals=list(range(len(resources))), ticktext=resources),
      title=title
    )
    return fig

  # Создание тепловых карт
  figPPE = create_heatmap(ppe_data, "Тепловая карта потребления первичной энергии")
  figKP = create_heatmap(kp_data, "Тепловая карта конечного потребления энергии")
  return [figPPE, figKP]

def heatMaps(tables):
  sectors = ['Сельское хозяйство', 'Промышленность', 'Строительство', 'Транспорт и связь', 'Сфера услуг', 'Население', 'Неэнергетические нужды'][::-1]
  resources = ['Уголь', 'Нефтепродукты', 'Природный газ', 'Электроэнергия', 'Тепловая энергия']
  res=[]
  for d in tables:
    year=d.iloc[0,10]
    d = d.iloc[2:].reset_index(drop=True)
    value = [
      [round(d.iloc[20, 2], 2), round(d.iloc[20, 4], 2), round(d.iloc[20, 5], 2), round(d.iloc[20, 9], 2), round(d.iloc[20, 10], 2)],
      [round(d.iloc[21, 2], 2), round(d.iloc[21, 4], 2), round(d.iloc[21, 5], 2), round(d.iloc[21, 9], 2), round(d.iloc[21, 10], 2)],
      [round(d.iloc[27, 2], 2), round(d.iloc[27, 4], 2), round(d.iloc[27, 5], 2), round(d.iloc[27, 9], 2), round(d.iloc[27, 10], 2)],
      [round(d.iloc[28, 2], 2), round(d.iloc[28, 4], 2), round(d.iloc[28, 5], 2), round(d.iloc[28, 9], 2), round(d.iloc[28, 10], 2)],
      [round(d.iloc[33, 2], 2), round(d.iloc[33, 4], 2), round(d.iloc[33, 5], 2), round(d.iloc[33, 9], 2), round(d.iloc[33, 10], 2)],
      [round(d.iloc[34, 2], 2), round(d.iloc[34, 4], 2), round(d.iloc[34, 5], 2), round(d.iloc[34, 9], 2), round(d.iloc[34, 10], 2)],
      [round(d.iloc[35, 2], 2), round(d.iloc[35, 4], 2), round(d.iloc[35, 5], 2), round(d.iloc[35, 9], 2), round(d.iloc[35, 10], 2)]][::-1]
    value=np.array(value)
    fig = go.Figure(data=go.Heatmap(
      z=value,
      x=resources,
      y=sectors,
      text=value,  # Текстовые метки
      texttemplate="%{text}",  # Отображение значений в ячейках
      colorscale='YlGnBu',
      colorbar=dict(title="Потребление")
    ))
    fig.update_layout(
      title={
          'text':year,
          'x':0.5,
          'xanchor': 'center',
          'yanchor': 'top'
        },
      xaxis_title='Источники энергии',
      yaxis_title='Секторы экономики',
    )
    res.append(fig)
  return res

def streamGraph(tables):
  sources = ['Расход ТЭР на производство ЭЭ', 'Расход ТЭР на производство ТЭ', 'Расход на преобразование топлива', 'Потери ЭЭ в сетях', 'Потери ТЭ в сетях', 'Потери на СП']
  data = {source: [] for source in sources}
  years=[]
  for table in tables:
    year=table.iloc[0,10]
    years.append(int(year[:4]))
    table = table.iloc[2:].reset_index(drop=True)
    if isnan(table.iloc[8,11]):
      data['Расход ТЭР на производство ЭЭ'].append(0)
    else:
      data['Расход ТЭР на производство ЭЭ'].append(abs(int(table.iloc[8,11])))
    if isnan(table.iloc[9,11]):
      data['Расход ТЭР на производство ТЭ'].append(0)
    else:
      data['Расход ТЭР на производство ТЭ'].append(abs(int(table.iloc[9,11])))
    if isnan(table.iloc[13,11]):
      data['Расход на преобразование топлива'].append(0)
    else:
      data['Расход на преобразование топлива'].append(abs(int(table.iloc[13,11])))
    if isnan(table.iloc[18,9]):
      data['Потери ЭЭ в сетях'].append(0)
    else:
      data['Потери ЭЭ в сетях'].append(abs(int(table.iloc[18,9])))
    if isnan(table.iloc[18,10]):
      data['Потери ТЭ в сетях'].append(0)
    else:
      data['Потери ТЭ в сетях'].append(abs(int(table.iloc[18,10])))
    if isnan(table.iloc[17,11]):
      data['Потери на СП'].append(0)
    else:
      data['Потери на СП'].append(abs(int(table.iloc[17,11])))

  data={**{'Year': years},**data}
  df=pd.DataFrame(data)
  fig=go.Figure()
  for source in sources:
    fig.add_trace(go.Scatter(
        x=df['Year'],
        y=df[source],
        mode='lines',
        stackgroup='one',  # Стек с накоплением
        name=source
    ))

  # Настройка графика
  fig.update_layout(
      title={
          'text':'Динамика эффективности сектора трансформации ТЭК региона',
          'x':0.5,
          'xanchor': 'center',
          'yanchor': 'top'
        },
      xaxis_title="Годы",
      yaxis_title="Значения, т.у.т.",
      legend_title="Категории",
      template="plotly_white",
      yaxis=dict(
          tickformat="d"  # Используем формат целого числа
      )
  )
  return [fig]

def radarDiagram(tables):
  VRP_MO=[4206.506,4290.2613,4644.635,5196.1364,5406.0765,6809.9511,7720.8426]#2016 2017 2018 2019 2020 2021 2022
  years=[]
  VRP=[]
  P1=[]
  P2=[]
  P3=[]
  P4=[]
  P5=[]
  P6=[]
  for d in tables:
    year=int(d.iloc[0,10][:4])
    years.append(year)
    VRP.append(VRP_MO[year-2016])
    d = d.iloc[4:].reset_index(drop=True)
    P1.append(1/abs(round(d.iloc[4,11]/VRP_MO[year-2016],3)))
    P2.append(abs(round(d.iloc[17,11]/d.iloc[4,11],3)))
    P3.append(abs(round( (d.iloc[6,9]+d.iloc[8,10])/(abs(d.iloc[6,2]+d.iloc[6,3]+d.iloc[6,4]+d.iloc[6,5]+d.iloc[6,6]+d.iloc[6,7]+d.iloc[6,8]+d.iloc[6,10])+abs(d.iloc[8,2]+d.iloc[8,3]+d.iloc[8,4]+d.iloc[8,5]+d.iloc[8,6]+d.iloc[8,7]+d.iloc[8,8]+d.iloc[8,9])),3)))
    P4.append(abs(round(d.iloc[9,10]/abs(d.iloc[9,2]+d.iloc[9,3]+d.iloc[9,4]+d.iloc[9,5]+d.iloc[9,6]+d.iloc[9,7]+d.iloc[9,8]+d.iloc[9,9]),3)))
    P5.append(1/abs(round(d.iloc[16,9]/d.iloc[17,9],3)))
    P6.append(1/abs(round(d.iloc[16,10]/d.iloc[17,10],3)))
  U1=[]
  U2=[]
  U3=[]
  U4=[]
  U5=[]
  U6=[]
  for i in range(1,len(years)): # За базовый берем первый год (2016г)
    U1.append(round(P1[i]/P1[0],2))
    U2.append(round(P2[i]/P2[0],2))
    U3.append(round(P3[i]/P3[0],2))
    U4.append(round(P4[i]/P4[0],2))
    U5.append(round(P5[i]/P5[0],2))
    U6.append(round(P6[i]/P6[0],2))

  years=[str(y) for y in years][1:]
  parameters=['U1', 'U2', 'U3', 'U4', 'U5', 'U6']
  data= {
    'U1':U1,
    'U2':U2,
    'U3':U3,
    'U4':U4,
    'U5':U5,
    'U6':U6,
  }
  # Преобразуем данные в удобный формат
  values_by_year = list(zip(*[data[param] for param in parameters]))

  # Создание фигуры
  fig1 = go.Figure()

  # Добавляем данные для каждого года
  for i, year in enumerate(years):
    values = values_by_year[i]
    fig1.add_trace(go.Scatterpolar(
        r=list(values) + [values[0]],  # Преобразуем кортеж в список и замыкаем круг
        theta=parameters + [parameters[0]],  # Замыкаем круг
        name=year
    ))

  # Настройка графика
  fig1.update_layout(
    polar=dict(
      radialaxis=dict(
          visible=True,
          range=[0, max(max(values_by_year)) + 0.5]  # Диапазон значений оси
      )
    ),
    title={
      'text':'Радарная диаграмма параметров по годам, 2016 год базовый',
      'x':0.5,
      'xanchor': 'center',
      'yanchor': 'top'
    },
    showlegend=True
  )

  # Создаем субплоты
  fig2 = make_subplots(
    rows=2, cols=3,  # Сетка 2x3
    specs=[[{'type': 'polar'}] * 3, [{'type': 'polar'}] * 3],  # Все графики полярные
    subplot_titles=years  # Названия графиков — годы
  )

  # Добавляем данные для каждого года
  for i, year in enumerate(years):
      values = values_by_year[i]
      row = i // 3 + 1  # Рассчитываем строку (целая часть от деления)
      col = i % 3 + 1   # Рассчитываем колонку (остаток от деления)
      fig2.add_trace(go.Scatterpolar(
          r=list(values) + [values[0]],  # Замыкаем круг
          theta=parameters + [parameters[0]],  # Замыкаем круг
          fill='toself',  # Закрашиваем область
          name=f"Year {year}",
          showlegend=False
      ), row=row, col=col)

  # Настройка графика
  fig2.update_layout(
      title={
      'text':'Радарные диаграммы по годам, 2016 год базовый',
      'x':0.5,
      'xanchor': 'center',
      'yanchor': 'top'
      },
      height=800,  # Высота всего графика
  )
  return[fig1,fig2]

def energyEfficiency(baseTable,currentTable):
  d=baseTable
  baseYear=d.iloc[0,10]
  baseYearInt=int(baseYear[:4])
  d = d.iloc[4:].reset_index(drop=True)
  b_p=[
    abs(round(d.iloc[4,11]/VRP_MO[baseYearInt-2016],3)),
    abs(round(d.iloc[17,11]/d.iloc[4,11],3)),
    abs(round( (d.iloc[6,9]+d.iloc[8,10])/(abs(d.iloc[6,2]+d.iloc[6,3]+d.iloc[6,4]+d.iloc[6,5]+d.iloc[6,6]+d.iloc[6,7]+d.iloc[6,8]+d.iloc[6,10])+abs(d.iloc[8,2]+d.iloc[8,3]+d.iloc[8,4]+d.iloc[8,5]+d.iloc[8,6]+d.iloc[8,7]+d.iloc[8,8]+d.iloc[8,9])),3)),
    abs(round(d.iloc[9,10]/abs(d.iloc[9,2]+d.iloc[9,3]+d.iloc[9,4]+d.iloc[9,5]+d.iloc[9,6]+d.iloc[9,7]+d.iloc[9,8]+d.iloc[9,9]),3)),
    abs(round(d.iloc[16,9]/d.iloc[17,9],3)),
    abs(round(d.iloc[16,10]/d.iloc[17,10],3))
  ]
  d=currentTable
  currentYear=d.iloc[0,10]
  currentYearInt=int(currentYear[:4])
  d = d.iloc[4:].reset_index(drop=True)
  p=[
    abs(round(d.iloc[4,11]/VRP_MO[currentYearInt-2016],3)),
    abs(round(d.iloc[17,11]/d.iloc[4,11],3)),
    abs(round( (d.iloc[6,9]+d.iloc[8,10])/(abs(d.iloc[6,2]+d.iloc[6,3]+d.iloc[6,4]+d.iloc[6,5]+d.iloc[6,6]+d.iloc[6,7]+d.iloc[6,8]+d.iloc[6,10])+abs(d.iloc[8,2]+d.iloc[8,3]+d.iloc[8,4]+d.iloc[8,5]+d.iloc[8,6]+d.iloc[8,7]+d.iloc[8,8]+d.iloc[8,9])),3)),
    abs(round(d.iloc[9,10]/abs(d.iloc[9,2]+d.iloc[9,3]+d.iloc[9,4]+d.iloc[9,5]+d.iloc[9,6]+d.iloc[9,7]+d.iloc[9,8]+d.iloc[9,9]),3)),
    abs(round(d.iloc[16,9]/d.iloc[17,9],3)),
    abs(round(d.iloc[16,10]/d.iloc[17,10],3))
  ]
  # Согласование «чем больше, тем лучше»
  b_p[0]=round(1/b_p[0],3)
  b_p[4]=round(1/b_p[4],3)
  b_p[5]=round(1/b_p[5],3)
  p[0]=round(1/p[0],3)
  p[4]=round(1/p[4],3)
  p[5]=round(1/p[5],3)
  # Определение значения функции полезности
  u=[] 
  for i in range(6):
    u.append(round(p[i]/b_p[i], 2))
  # уровень значимости показателей
  #r = [0.17, 0.17, 0.17, 0.17, 0.16, 0.16]
  r = [0.2, 0.2, 0.1, 0.2, 0.1, 0.2]
  # узловые точки в классификаторе
  g = [0.1, 0.3, 0.5, 0.7, 0.9]
  # Классификация отдельных значений
  M = [[0.55, 0.65, 0.75], [0.65, 0.75, 0.85, 0.95], [0.85, 0.95, 1.05, 1.15], [1.05, 1.15, 1.25, 1.35], [1.25, 1.35, 1.45]]
  # Классификация уровней показателей
  lvlX = [[0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0]]
  for i in range(6):
    b = [] # какие значения лингв переменной может принимать (0 - Плохо, 1 - Есть ухудшение, 2 - Не хуже, 3 - Есть улучшение, 4 - Хорошо)
    w = [] # интервал значения показателя
    for j in range(5):
      for h in range(len(M[j]) - 1):
        if M[j][h] <= u[i] <= M[j][h+1]:
          w = [M[j][h], M[j][h+1]]
          b.append(j)
          break
    if len(b) == 2:
      if u[i]==w[0]:
        lvlX[i][b[0]]=0.95
        lvlX[i][b[1]]=0.05
      elif u[i]==w[1]:
        lvlX[i][b[0]]=0.05
        lvlX[i][b[1]]=0.95
      else:
        lvlX[i][b[0]] = round(10*(w[1] - u[i]), 2)
        lvlX[i][b[1]] = round(1 - lvlX[i][b[0]],2)
    elif len(b) == 1:
      lvlX[i][b[0]] = 1
    elif len(b) == 0:
      lvlX[i][4] = 1
  # считаем Динамику изменения энергопотребления в регионе
  D = 0  
  for i in range(6):
    s = 0
    for j in range(5):
      s += g[j] * lvlX[i][j]
    D += r[i] * s
  MD = [[0, 0.15, 0.25], [0.15, 0.25, 0.35, 0.45], [0.35, 0.45, 0.55, 0.65], [0.55, 0.65, 0.75, 0.85], [0.75, 0.85, 1]]
  lvlD = [0, 0, 0, 0, 0]
  b = [] # какие значения лингв переменной может принимать (0 - Отрицателньая, 1 - Есть ухудшение, 2 - Без изменения, 3 - Есть улучшение, 4 - Положительная)
  w = [] # интервал значения показателя
  for j in range(5):
    for h in range(len(M[j]) - 1):
      if MD[j][h] <= D <= MD[j][h+1]:
        w = [M[j][h],MD[j][h+1]]
        b.append(j)
        break
  if len(b) == 2:
    if D==w[0]:
      lvlD[b[0]]=0.9
      lvlD[b[1]]=0.05
    elif D==w[1]:
      lvlD[b[0]]=0.05
      lvlD[b[1]]=0.95
    else:
      lvlD[b[0]] = round(10*(w[1] - D), 2)
      lvlD[b[1]] = round(1 - lvlD[b[0]],2)
  elif len(b) == 1:
    lvlD[b[0]] = 1
  elif len(b) == 0 :
    lvlD[4] = 1

  return lvlD, D, lvlX

def prepare_data_for_display(lvlD, D, lvlX):
   # Логика генерации текста отчета
    if lvlD[0] == 1 or (lvlD[0] == 0.5 and lvlD[1] == 0.5) or (lvlD[0] > 0 and lvlD[1] > 0 and lvlD[0] > lvlD[1]):
      lpD = 'Отрицательная'
    elif lvlD[1] == 1 or (lvlD[1] == 0.5 and lvlD[2] == 0.5) or (lvlD[1] > 0 and lvlD[2] > 0 and lvlD[1] > lvlD[2]) or (lvlD[0] > 0 and lvlD[1] > 0 and lvlD[0] < lvlD[1]):
      lpD = 'Есть ухудшение'
    elif lvlD[2] == 1 or (lvlD[2] == 0.5 and lvlD[3] == 0.5) or (lvlD[2] > 0 and lvlD[3] > 0 and lvlD[2] > lvlD[3]) or (lvlD[1] > 0 and lvlD[2] > 0 and lvlD[1] < lvlD[2]):
      lpD = 'Без изменения'
    elif lvlD[3] == 1 or (lvlD[3] == 0.5 and lvlD[4] == 0.5) or (lvlD[3] > 0 and lvlD[4] > 0 and lvlD[3] > lvlD[4]) or (lvlD[2] > 0 and lvlD[3] > 0 and lvlD[2] < lvlD[3]):
      lpD = 'Есть улучшение'
    elif (lvlD[3] > 0 and lvlD[4] > 0 and lvlD[4] > lvlD[3]) or lvlD[4] == 1:
      lpD = 'Положительная'

    report_text = f"""
    <h2>Значение лингвистической переменной D: {D}</h2>
    <p>Динамика эффективности использования ТЭР в регионе: {lpD}</p>
    """
    indicators = [
      "Энергоемкость ВРП (P1)", 
      "Коэффициент полезного использования ТЭР в конечном потр. (P2)",
      "Коэффициент полезного использования ТЭР на электростанциях (P3)",
      "Коэффициент полезного использования ТЭР котельными (P4)",
      "Доля потерь в сетях в конечном потр. электрической энергии (P5)",
      "Доля потерь в сетях в конечном потр. тепловой энергии (P6)"
    ]
    for i in range(6):
      if lvlX[i][0] == 1 or (lvlX[i][0] == 0.5 and lvlX[i][1] == 0.5) or (lvlX[i][0] > 0 and lvlX[i][1] > 0 and lvlX[i][0] > lvlX[i][1]):
        lpP = 'Плохо'
      elif lvlX[i][1] == 1 or (lvlX[i][1] == 0.5 and lvlX[i][2] == 0.5) or (lvlX[i][1] > 0 and lvlX[i][2] > 0 and lvlX[i][1] > lvlX[i][2]) or (lvlX[i][0] > 0 and lvlX[i][1] > 0 and lvlX[i][0] < lvlX[i][1]):
        lpP = 'Есть ухудшение'
      elif lvlX[i][2] == 1 or (lvlX[i][2] == 0.5 and lvlX[i][3] == 0.5) or (lvlX[i][2] > 0 and lvlX[i][3] > 0 and lvlX[i][2] > lvlX[i][3]) or (lvlX[i][1] > 0 and lvlX[i][2] > 0 and lvlX[i][1] < lvlX[i][2]):
        lpP = 'Не хуже'
      elif lvlX[i][3] == 1 or (lvlX[i][3] == 0.5 and lvlX[i][4] == 0.5) or (lvlX[i][3] > 0 and lvlX[i][4] > 0 and lvlX[i][3] > lvlX[i][4]) or (lvlX[i][2] > 0 and lvlX[i][3] > 0 and lvlX[i][2] < lvlX[i][3]):
        lpP = 'Есть улучшение'
      elif (lvlX[i][3] > 0 and lvlX[i][4] > 0 and lvlX[i][4] > lvlX[i][3]) or lvlX[i][4] == 1:
        lpP = 'Хорошо'
      report_text += f"<p>Уровень полезности показателя {indicators[i]}: {lpP}</p>"
    report_text += "<h3>Лингвистическая переменная оценки уровня полезности показателей:</h3>"
    report_text += "<table border='1'><tr><th>Индикатор</th><th>Плохо</th><th>Есть ухудшение</th><th>Не хуже</th><th>Есть улучшение</th><th>Хорошо</th></tr>"
    for i in range(6):
      report_text += f"""
      <tr>
        <td>P{i+1}</td>
        <td>{lvlX[i][0]}</td>
        <td>{lvlX[i][1]}</td>
        <td>{lvlX[i][2]}</td>
        <td>{lvlX[i][3]}</td>
        <td>{lvlX[i][4]}</td>
      </tr>
      """
    report_text += "</table>"
    return report_text

def allowed_file(filename):
  return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
  files = os.listdir(app.config['UPLOAD_FOLDER'])
  return render_template('index.html', files=files)

@app.route('/download-template')
def download_template():
  try:
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], 'template.xlsm', as_attachment=True)
  except FileNotFoundError:
    return "Шаблон не найден", 404

@app.route('/upload', methods=['POST'])
def upload_files():
  if 'files' not in request.files:
    return redirect(request.url)
  
  files = request.files.getlist('files')

  if not files or any(f.filename == '' for f in files):
    return redirect(request.url)

  for file in files:
    if file and allowed_file(file.filename):
      filename = secure_filename(file.filename)
      filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
      file.save(filepath)

  return redirect(url_for('index'))

@app.route('/delete_files')
def delete_files():
  clear_upload_folder()
  return redirect(url_for('index'))

@app.route('/display-report', methods=['POST'])
def display_report():
  try:
    base_file = request.form.get('base_file')
    current_file = request.form.get('current_file')

    if not base_file or not current_file:
      return redirect(url_for('index'))

    base_filepath = os.path.join(app.config['UPLOAD_FOLDER'], base_file)
    current_filepath = os.path.join(app.config['UPLOAD_FOLDER'], current_file)

    if not os.path.exists(base_filepath) or not os.path.exists(current_filepath):
      return "Файл не найден", 404

    fig_html = analyze_excel_files(base_filepath, current_filepath)

    return render_template('index.html', files=os.listdir(app.config['UPLOAD_FOLDER']), fig_html=fig_html)
  except Exception as e:
    app.logger.error(f"Error in /display-report: {str(e)}")
    return "Internal Server Error", 500
@app.route('/download-report', methods=['POST'])
def download_report():
  base_file = request.form.get('base_file')
  current_file = request.form.get('current_file')

  if not base_file or not current_file:
    return redirect(url_for('index'))

  base_filepath = os.path.join(app.config['UPLOAD_FOLDER'], base_file)
  current_filepath = os.path.join(app.config['UPLOAD_FOLDER'], current_file)

  if not os.path.exists(base_filepath) or not os.path.exists(current_filepath):
    return "Файл не найден", 404

  pdf_path = os.path.join(app.config['DOWNLOAD_FOLDER'], 'report.pdf')
  #doc.SaveToFile(pdf_path, FileFormat.PDF)
  return send_from_directory(app.config['DOWNLOAD_FOLDER'], 'report.pdf', as_attachment=True)
  
def analyze_excel_files(base_filepath, current_filepath):
  files = os.listdir(app.config['UPLOAD_FOLDER'])
  tables = [pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], f),header=None) for f in files if allowed_file(f)] 
  tables2 = [pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], f),header=None,sheet_name=None) for f in files if allowed_file(f)] 
  baseTable = pd.read_excel(base_filepath, header=None)
  currentTable = pd.read_excel(current_filepath, header=None)

  figSankey = [sankey(t) for t in tables]
  figSankey2 = [sankey2(t) for t in tables]
  figPiePPE = [piePPE(t) for t in tables]
  figPieKP = [pieKP(t) for t in tables]
  figBar = barChart(tables)
  figPar=parcoordsChart(tables, tables2)
  figHeatMapsGeneral = heatMapsGeneral(tables)
  figHeatMaps = heatMaps(tables)
  figStream = streamGraph(tables)
  figRadar = radarDiagram(tables)
  lvlD, D, lvlX =energyEfficiency(baseTable, currentTable)
  report_text = prepare_data_for_display(lvlD,D,lvlX)
  figs = figSankey+figSankey2+figPar+figHeatMapsGeneral+figHeatMaps+figStream+figRadar+figPiePPE+figPieKP+figBar
  fig_html = ''.join([pio.to_html(fig, full_html=False) for fig in figs])
  # Объединяем отчет и графики
  full_content = f"<div>{report_text}</div>{fig_html}"
  return full_content

# Функция для очистки папки uploads
def clear_upload_folder():
  folder = app.config['UPLOAD_FOLDER']
  if os.path.exists(folder):
    for filename in os.listdir(folder):
      file_path = os.path.join(folder, filename)
      try:
        if os.path.isfile(file_path) or os.path.islink(file_path):
          os.unlink(file_path)  # Удаляем файл или ссылку
        elif os.path.isdir(file_path):
          shutil.rmtree(file_path)  # Удаляем папку и её содержимое
      except Exception as e:
        print(f"Не удалось удалить {file_path}. Причина: {e}")

if __name__ == '__main__':
  clear_upload_folder()  # Очистка папки перед запуском
  from waitress import serve
  serve(app, host="0.0.0.0", port=8080)
  # app.run(debug=True)
