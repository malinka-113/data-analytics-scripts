#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install sqlparse openpyxl pandas


# In[7]:


import re
import sqlparse
from sqlparse.sql import IdentifierList, Identifier, Function, Where
from sqlparse.tokens import Keyword, DML, Punctuation
import pandas as pd
from pathlib import Path


class SQLParser:
    """Парсер SQL-запросов для извлечения таблиц и атрибутов"""

    def __init__(self, sql_text: str):
        self.sql_text = sql_text
        self.parsed = sqlparse.parse(sql_text)[0]
        self.tables = {}  # alias -> full_table_name
        self.attributes = []  # (table_alias_or_name, column_name)

    # ============================================================
    # 1. Извлечение таблиц (FROM + JOIN)
    # ============================================================
    def extract_tables(self) -> dict:
        """
        Извлекает все таблицы из FROM и JOIN с их алиасами.
        Возвращает словарь: {alias: full_table_name}
        """
        tables = {}

        # Паттерн для FROM и JOIN
        pattern = re.compile(
            r'(?:FROM|JOIN)\s+([A-Za-z0-9_.]+)'  # имя таблицы
            r'(?:\s+(?:AS\s+)?([A-Za-z0-9_]+))?',  # опциональный алиас
            re.IGNORECASE
        )

        for match in pattern.finditer(self.sql_text):
            table_name = match.group(1).strip()
            alias = match.group(2).strip().lower() if match.group(2) else table_name.lower()
            tables[alias] = table_name

        self.tables = tables
        return tables

    # ============================================================
    # 2. Извлечение атрибутов из SELECT
    # ============================================================
    def extract_attributes(self) -> list:
        """
        Извлекает атрибуты из SELECT с привязкой к таблицам.
        Возвращает список кортежей: (table_name, column_name)
        """
        attributes = []

        # Находим блок SELECT ... FROM
        select_match = re.search(
            r'SELECT\s+(.*?)\s+FROM\s+',
            self.sql_text,
            re.IGNORECASE | re.DOTALL
        )
        if not select_match:
            return attributes

        select_block = select_match.group(1)

        # Разбиваем SELECT на отдельные выражения по запятым
        # (учитываем вложенные скобки)
        expressions = self._split_select_expressions(select_block)

        for expr in expressions:
            expr = expr.strip()
            if not expr:
                continue

            # Пропускаем константы и NULL без алиаса таблицы
            if re.match(r"^(NULL|'[^']*'|[\d.]+)\s*(?:AS\s+\w+)?$", expr, re.IGNORECASE):
                alias_match = re.search(r'AS\s+(\w+)', expr, re.IGNORECASE)
                col_name = alias_match.group(1) if alias_match else expr
                attributes.append(('', col_name))
                continue

            # Паттерн: table_alias.column_name [AS alias]
            col_match = re.match(
                r'([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)',
                expr
            )
            if col_match:
                table_alias = col_match.group(1).lower()
                column_name = col_match.group(2)
                full_table = self.tables.get(table_alias, table_alias)
                attributes.append((full_table, column_name))
                continue

            # Паттерн: функция(table_alias.column_name)
            func_match = re.search(
                r'\w+\(([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)',
                expr
            )
            if func_match:
                table_alias = func_match.group(1)
                column_name = func_match.group(2)
                full_table = self.tables.get(table_alias, table_alias)
                attributes.append((full_table, column_name))
                continue

            # Паттерн: NVL(a.col1, b.col2) — несколько колонок
            nvl_matches = re.findall(
                r'([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)',
                expr
            )
            if nvl_matches:
                for table_alias, column_name in nvl_matches:
                    full_table = self.tables.get(table_alias, table_alias)
                    attributes.append((full_table, column_name))
                continue

            # Просто колонка без алиаса таблицы
            simple_match = re.match(r'^([A-Za-z_]\w*)\s*(?:AS\s+\w+)?$', expr, re.IGNORECASE)
            if simple_match:
                col_name = simple_match.group(1)
                attributes.append(('', col_name))
                continue

            # Алиас для константы/выражения
            alias_match = re.search(r'AS\s+(\w+)', expr, re.IGNORECASE)
            if alias_match:
                attributes.append(('', alias_match.group(1)))

        self.attributes = attributes
        return attributes

    # ============================================================
    # 3. Вспомогательная функция: разбиение SELECT на выражения
    # ============================================================
    def _split_select_expressions(self, select_block: str) -> list:
        """Разбивает SELECT на выражения, учитывая вложенные скобки"""
        expressions = []
        current = []
        depth = 0

        for char in select_block:
            if char == '(':
                depth += 1
                current.append(char)
            elif char == ')':
                depth -= 1
                current.append(char)
            elif char == ',' and depth == 0:
                expressions.append(''.join(current))
                current = []
            else:
                current.append(char)

        if current:
            expressions.append(''.join(current))

        return expressions

    # ============================================================
    # 4. Выгрузка в Excel
    # ============================================================
    def save_to_excel(self, output_path: str):
        """Сохраняет результат парсинга в Excel"""
        # Таблицы
        tables_df = pd.DataFrame(
            list(self.tables.items()),
            columns=['alias', 'table_name']
        )
        tables_df = tables_df.drop_duplicates()

        # Атрибуты
        attrs_df = pd.DataFrame(
            self.attributes,
            columns=['table_name', 'attribute_name']
        )
        # Убираем дубликаты
        attrs_df = attrs_df.drop_duplicates()

        # Сохраняем в Excel на разных листах
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            tables_df.to_excel(writer, sheet_name='Таблицы', index=False)
            attrs_df.to_excel(writer, sheet_name='Атрибуты', index=False)

        print(f"✅ Результат сохранён в {output_path}")
        print(f"   Таблиц: {len(tables_df)}, Атрибутов: {len(attrs_df)}")

        return tables_df, attrs_df


# ============================================================
# Пример использования
# ============================================================
if __name__ == "__main__":
    # Чтение SQL из файла
    sql_file = r"C:\Users\EPavlova\Documents\ИЗ ВРМ\СУБО 131 (AXDP)\ГЭП ОВЮЛ\axdp - маппинги овюл\SQL-запросы\query.sql"
    with open(sql_file, 'r', encoding='utf-8') as f:
        sql_query = f.read()

    parser = SQLParser(sql_query)
    parser.extract_tables()
    parser.extract_attributes()
    parser.save_to_excel("sql_parsed_result2.xlsx")


    # Извлечение таблиц
    print("=" * 60)
    print(" ТАБЛИЦЫ")
    print("=" * 60)
    tables = parser.extract_tables()
    for alias, name in tables.items():
        print(f"  {alias:10} → {name}")

    # Извлечение атрибутов
    print("\n" + "=" * 60)
    print(" АТРИБУТЫ")
    print("=" * 60)
    attributes = parser.extract_attributes()
    for table, attr in attributes:
        print(f"  {table:45} | {attr}")

    # Выгрузка в Excel
    parser.save_to_excel("sql_parsed_result.xlsx")


# In[ ]:




