# ETL-процесс и визуализация по данным о фильмах.

## Описание проекта
Этот проект реализует полный ETL-конвейер для анализа данных о фильмах.  
Он состоит из двух подпроектов в Visual Studio 2022:  
**C#-подпроект** — загружает Excel-файлы в SQL Server.  
**Python-подпроект** — анализирует данные, выполняет агрегации и визуализирует результаты.


## Структура решения в Visual Studio
<pre> ```
/Solution
├─ /ExcelLoaderWithEPPlus # C#-подпроект
│ └─ Program.cs (ExcelLoaderWithEPPlus.cs)
│
├─ /ETL_Movies_Analysis # Python-подпроект
│ └─ etl_movies.py
│
└─ README.md # Документация (этот файл)
``` </pre>

## Подпроект 1: Загрузка данных (C#)
### Требования
- .NET 8.0 SDK
- NuGet-пакеты:
  - EPPlus
  - Microsoft.Data.SqlClient
	
### Описание
- Использует **EPPlus** для чтения Excel-файлов.
- Создаёт базу данных **ExcelDataDB** (если её нет).
- Для каждого Excel-файла создаёт таблицу в БД.
- Загружает данные через **SqlBulkCopy**.

### Настройка
1️) Путь к папке с Excel-файлами:
	<pre> ``` string folderPath = @"C:\Users\yersu\Desktop\Для интервью\Тех Задание"; ``` </pre>
2) Строка подключения:
	 <pre> ``` string connectionStringMaster = "Server=localhost;Database=master;Trusted_Connection=True;TrustServerCertificate=True;"; ``` </pre>

### Запуск
- Откройте проект в Visual Studio.
- Нажимаете F5 для запуска.
- Вывод: 
	Обрабатывается файл: content_film_work
    Создана таблица: content_film_work
    Данные загружены в: content_film_work
    ...		
    Все Excel-файлы обработаны!

## Подпроект 2: Анализ и визуализация (Python)
### Требования
- Python 3.9+
- Библиотеки Python:
  - pandas
  - sqlalchemy
  - matplotlib
  - pyodbc
  - seaborn
- Драйвер для подключения к SQL Server: ODBC Driver 17 for SQL Server (или более новая версия)
- Доступ к базе данных SQL Server (по умолчанию используется Windows Authentication).
	
### Описание
Скрипт etl_movies.py:
- Загружает данные из БД ExcelDataDB.
- Приводит даты (creation_date, created) к datetime.
- Убирает дубликаты, приводит рейтинг к float.
- Строит агрегаты:
  - Кол-во фильмов по годам
  - Средняя оценка по годам
  - Кол-во фильмов по жанрам
  - Распределение ролей
  - Топ-10 актёров
  - Гистограмма рейтингов
- Распределение по типу
  - Сохраняет агрегаты в CSV.
  - Создаёт PNG-графики и PDF-отчёт с графиками.

### Настройка
1. Установи зависимости:
   <pre> ``` pip install pandas sqlalchemy matplotlib pyodbc seaborn ``` </pre>
2. Строка подключения к БД:
   <pre> ``` connection_string = "mssql+pyodbc://localhost/ExcelDataDB?driver=ODBC+Driver+17+for+SQL+Server" ``` </pre>

### Запуск
1. Запустите скрипт в терминале:
  python etl_movies.py
2. Результат:
3. Создаются CSV-файлы:
  - films_by_year.csv
  - avg_rating_by_year.csv
  - genre_stats.csv
  - role_distribution.csv
  - actor_top10.csv
  - type_distribution.csv
- Создаются PNG-графики:
  - films_by_year.png
  - avg_rating_by_year.png
  - genre_stats.png
  - role_distribution.png
  - actor_top10.png
  - type_distribution.png


# Ключевые моменты ETL

## 1. Загрузка (C#)

Извлекает данные из Excel.
Создаёт таблицы в БД.
Загружает данные.

## 2. Трансформация (Python)

Приводит даты (creation_date, created) к datetime.
Не заполняет creation_date датой создания записи (created), чтобы сохранить корректность данных.
Преобразует рейтинг в float.
Убирает дубликаты.

## 3. Агрегация и визуализация (Python)

Агрегирует данные по разным срезам.
Строит графики с помощью Seaborn и Matplotlib.
Сохраняет результаты в CSV, PNG.

# Вывод 

- Проект готов к загрузке и анализу фильмов.
- Графики и отчёты сохраняются автоматически.
- Удобно использовать как базу для любых дальнейших аналитических задач.