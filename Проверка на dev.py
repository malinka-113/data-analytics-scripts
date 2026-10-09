#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().system('pip install pandas psycopg2-binary openpyxl sqlalchemy')


# In[2]:


get_ipython().system('pip install psycopg2-binary')


# In[1]:


import pandas as pd
from sqlalchemy import create_engine, text
import os


# In[6]:


# --- Путь к Excel файлу ---
EXCEL_FILE = "C:/Users/EPavlova/Documents/ИЗ ВРМ/codereview.xlsx"          # Полный путь или просто имя, если файл в той же папке

# --- Подключение к PostgreSQL ---
DB_HOST = "localhost"
DB_PORT = "5432"
DB_NAME = "postgres"         
DB_USER = "postgres"         
DB_PASSWORD = "**********"     

# --- Имя таблицы ---
TABLE_NAME = "dev_codereview"

# --- Сопоставление колонок ---
COLUMN_MAPPING = {
    "ID": "id_check",              
    "Название проверки": "name_check",
    "Тип": "type",
    "Объект": "object",
    "Связанный объект": "related_object",
    "Описание": "description"
}


# In[7]:


# Проверяем файл
if not os.path.exists(EXCEL_FILE):
    print(f"❌ Файл '{EXCEL_FILE}' не найден!")
    print(f"📁 Текущая папка: {os.getcwd()}")
else:
    # Читаем Excel
    df = pd.read_excel(EXCEL_FILE, sheet_name=0, engine='openpyxl')
    print(f"✅ Прочитано строк: {len(df)}")
    print(f"📊 Колонки в Excel: {list(df.columns)}")

    # Переименовываем колонки
    df = df[list(COLUMN_MAPPING.keys())]
    df = df.rename(columns=COLUMN_MAPPING)

    # Приводим типы
    df['id_check'] = pd.to_numeric(df['id_check'], errors='coerce')
    for col in ['name_check', 'type', 'object', 'related_object', 'description']:
        if col in df.columns:
            df[col] = df[col].astype(str).replace('nan', None)
            # Обрезаем до 50 символов (чтобы не было ошибки VARCHAR(50))
            df[col] = df[col].str[:50]

    print(f"\n📊 Первые 3 строки для загрузки:")
    display(df.head(3))  # display работает только в Jupyter!


# In[8]:


# Подключаемся и загружаем
engine = create_engine(
    f'postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
)

# Проверяем соединение
with engine.connect() as conn:
    print("✅ Подключение к PostgreSQL успешно")

# Загружаем
df.to_sql(
    TABLE_NAME,
    con=engine,
    if_exists='append',
    index=False,
    method='multi',
    chunksize=1000
)

print(f"✅ Успешно загружено {len(df)} строк в таблицу '{TABLE_NAME}'")


# In[9]:


# Проверяем, что загрузилось
with engine.connect() as conn:
    result = conn.execute(text(f"SELECT COUNT(*) FROM {TABLE_NAME}"))
    total = result.fetchone()[0]
    print(f"📊 Всего записей в таблице: {total}")

    print("\n📋 Последние 5 записей:")
    result = conn.execute(text(f"SELECT * FROM {TABLE_NAME} ORDER BY id_check DESC LIMIT 5"))
    for row in result:
        print(row)


# In[ ]:




