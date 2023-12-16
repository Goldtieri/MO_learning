import math
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

# Загрузка данных из файла Excel
file_path = 'C:\\Univer\\lab1\\MOLAB2_1\\new_data.xlsx'
df = pd.read_excel(file_path)

newIndex = None
newUzel = 0
newKonec = 0

for index, row in df.iterrows():
    if row['class'] == 'new':
        newIndex = index
        newUzel = row['Nuzl']
        newKonec = row['Nk']


if newIndex == None:
    exit()

minN = 0
minDist = math.inf
minClass = ""

table = []
countByClasses = {'I': 0, 'II': 0, 'III': 0}
maxCount = 0

flag = True

for (index, row) in df.iterrows():
    if row['class'] == 'new':
        continue

    df.at[index, 'distance'] = math.sqrt((row['Nuzl'] - newUzel) ** 2 + (row['Nk'] - newKonec) ** 2)
    if df.at[index, 'distance'] < minDist:
        minDist = df.at[index, 'distance']

while True:
    R = int(input(f"Введите R не меньше чем {minDist}: "))

    if not (R < minDist):
        break

df.at[0, 'R'] = R
minDist = R + 1

for (index, row) in df.iterrows():
    if df.at[index, 'distance'] <= R and row['class'] != 'new':
        table.append([row['№'], row['class'], df.at[index, 'distance'], row['Nk'], row['Nuzl']])
        countByClasses[row['class']] += 1

        if countByClasses[row['class']] > maxCount:
            maxCount = countByClasses[row['class']]

        flag = False

conflict = False
count = 0
for key, value in countByClasses.items():
    if value == maxCount:
        count += 1

if count > 1:
    conflict = True

print("\nТочки в области R:")
print(table)

print("\nКоличество точек в области по классам:")
print(countByClasses)

if flag:
    exit()

minKonec = None
minUzel = None

for item in table:
    if countByClasses[item[1]] == maxCount and item[2] <= minDist:
        minDist = item[2]
        minN = item[0]
        minClass = item[1]
        minKonec = item[3]
        minUzel = item[4]

print("\nБлижайшая точка к новой:")
print("№:", minN, " Class:", minClass, " Distance:", minDist)

for index, row in df.iterrows():
    if index == 0:
        df.at[0, 'closest'] = minN
        if minClass == 'I':
            df.at[0, 'count1'] += 1
        elif minClass == 'II':
            df.at[0, 'count2'] += 1
        else:
            df.at[0, 'count3'] += 1
    if row['class'] == 'new':
        df.at[index, 'class'] = minClass

df.to_excel('C:\\Univer\\lab1\\MOLAB2_1\\new_data.xlsx', index=False)

plt.figure(figsize=(7, 10))  # Размер графика

# Создание словаря для соответствия классов и цветов
class_colors = {
    'I': 'r',  # Красный цвет для класса 1
    'II': 'g',  # Зеленый цвет для класса 2
    'III': 'b',  # Синий цвет для класса 3
    'new': 'y'
}

print("\nЦвета точек по классам: ", class_colors)

# Построение точек на графике
for index, row in df.iterrows():
    x = row['Nk']
    y = row['Nuzl']
    point_class = row['class']

    # Получите цвет для данного класса
    if x == newKonec and y == newUzel:
        point_class = 'new'
    point_color = class_colors.get(point_class, 'k')

    plt.scatter(x, y, color=point_color, label=point_class)

# Отображение окружности
circle = plt.Circle((newKonec, newUzel), R, color='black', fill=False)  # Создание окружности
plt.gca().add_patch(circle)  # Добавление окружности на график
if conflict:
    plt.plot([newKonec, minKonec], [newUzel, minUzel], color='black')
# Настройка графика
plt.xlabel('Координата X')
plt.ylabel('Координата Y')
plt.title('График с точками и цветами классов')

# Отображение графика
plt.axis('equal')
plt.show()
