import pandas as pd
import os

# Ввод пути к Excel-файлу
file_path = input("Введите путь к Excel-файлу: ")

# Загрузка таблицы
df = pd.read_excel(file_path)

df = df.rename(columns={df.columns[0]: 'Атрибут'})

# Преобразование в длинный формат
df_melted = df.melt(id_vars='Атрибут', var_name='Класс', value_name='Значение')

# Фильтрация только непустых значений
df_filtered = df_melted[df_melted['Значение'].notna() & (df_melted['Значение'] != '')]

# Печать результата в консоль
print("\nПреобразованные данные:\n")
print(df_filtered)

# Создание пути для сохранения
folder = os.path.dirname(file_path)
filename = os.path.splitext(os.path.basename(file_path))[0]
output_path = os.path.join(folder, f"{filename}_converted.xlsx")

# Сохранение результата
df_filtered.to_excel(output_path, index=False)
print(f"\nФайл сохранён: {output_path}")
