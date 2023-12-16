import pandas as pd
import math

# Загрузка данных из Excel файла
data = pd.read_excel('data.xlsx')

# Параметры парзентовского окна
R = 10  # Радиус окна

# Найти объект с классом 'new'
new_object = data[data['class'] == 'new'].iloc[0]

# Найти ближайший объект в пределах расстояния R
closest_object = None
closest_dist = math.inf

for index, obj in data.iterrows():
    if obj['class'] != 'new':
        distance = math.sqrt((obj['Nuzl'] - new_object['Nuzl'])**2 + (obj['Nk'] - new_object['Nk'])**2)
        if distance <= R and distance < closest_dist:
            closest_object = obj
            closest_dist = distance

# Изменить класс нового объекта и обновить информацию о ближайшем объекте
if closest_object is not None:
    new_object['class'] = closest_object['class']
    new_object['closest'] = closest_object['№']
    # Обновить количество объектов каждого класса
    if closest_object['class'] == 'I':
        data.loc[data['class'] == 'I', 'count1'] += 1
    elif closest_object['class'] == 'II':
        data.loc[data['class'] == 'II', 'count2'] += 1
    elif closest_object['class'] == 'III':
        data.loc[data['class'] == 'III', 'count3'] += 1

# Обновить информацию о новом объекте в исходных данных
data.loc[data['№'] == new_object['№']] = new_object

# Сохранить обновленные данные в Excel файл
data.to_excel('upData.xlsx', index=False)
