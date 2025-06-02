import pandas as pd
from sqlalchemy import create_engine
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages


engine = create_engine(
    "mssql+pyodbc://localhost/ExcelDataDB?"
    "driver=ODBC+Driver+17+for+SQL+Server&"
    "Trusted_Connection=yes&"
    "TrustServerCertificate=yes"
)

print("Шаг 1")
film_work = pd.read_sql('SELECT * FROM content_film_work', engine)
genre = pd.read_sql('SELECT * FROM content_genre', engine)
genre_film_work = pd.read_sql('SELECT * FROM content_genre_film_work', engine)
person = pd.read_sql('SELECT * FROM content_person', engine)
person_film_work = pd.read_sql('SELECT * FROM content_person_film_work', engine)

print("Шаг 1.5")
film_work['created'] = pd.to_datetime(film_work['created'], errors='coerce')
film_work['creation_date'] = pd.to_datetime(film_work['creation_date'], errors='coerce')

mask_na_creation_date = film_work['creation_date'].isna()
film_work.loc[mask_na_creation_date, 'creation_date'] = film_work.loc[mask_na_creation_date, 'created']


film_work['creation_date'] = film_work['creation_date'].dt.year

film_work['rating'] = pd.to_numeric(film_work['rating'], errors='coerce').fillna(0.0)

genre_film_work = genre_film_work.drop_duplicates(subset=['film_work_id', 'genre_id'])
person_film_work = person_film_work.drop_duplicates(subset=['film_work_id', 'person_id', 'role'])


print("Шаг 2")
film_genre_df = (film_work
                  .merge(genre_film_work, left_on='id', right_on='film_work_id', how='left')
                  .merge(genre, left_on='genre_id', right_on='id', suffixes=('', '_genre')))

film_person_df = (film_work
                   .merge(person_film_work, left_on='id', right_on='film_work_id', how='left')
                   .merge(person, left_on='person_id', right_on='id', suffixes=('', '_person')))



films_by_year = (film_work.groupby('creation_date')
                  .size().reset_index(name='count')
                  .sort_values('creation_date'))


avg_rating_by_year = (film_work.groupby('creation_date')['rating']
                       .mean().reset_index(name='avg_rating')
                       .sort_values('creation_date'))

genre_stats = (film_genre_df.groupby('name')
                .size().reset_index(name='count')
                .sort_values(by='count', ascending=False))

role_distribution = (person_film_work.groupby('role')
                      .size().reset_index(name='count')
                      .sort_values(by='count', ascending=False))

actors = film_person_df[film_person_df['role'] == 'actor']
actor_top10 = (actors.groupby('full_name')
                .size().reset_index(name='count')
                .sort_values(by='count', ascending=False)
                .head(10))

rating_histogram = film_work['rating'].dropna()

type_distribution = (film_work.groupby('type')
                      .size().reset_index(name='count')
                      .sort_values(by='count', ascending=False))


films_by_year.to_csv('data/films_by_year.csv', index=False)
avg_rating_by_year.to_csv('data/avg_rating_by_year.csv', index=False)
genre_stats.to_csv('data/genre_stats.csv', index=False)
role_distribution.to_csv('data/role_distribution.csv', index=False)
actor_top10.to_csv('data/actor_top10.csv', index=False)
type_distribution.to_csv('data/type_distribution.csv', index=False)


sns.set(style="whitegrid")
print("Шаг 3")

plt.figure(figsize=(10,5))
sns.barplot(x='creation_date', y='count', data=films_by_year, palette='viridis')
plt.title('Количество фильмов по годам')
plt.xlabel('Год')
plt.ylabel('Количество')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('data/films_by_year.png')
plt.close()


plt.figure(figsize=(10,5))
sns.lineplot(x='creation_date', y='avg_rating', data=avg_rating_by_year, marker='o', color='teal')
plt.title('Средняя оценка фильмов по годам')
plt.xlabel('Год')
plt.ylabel('Средняя оценка')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig('data/avg_rating_by_year.png')
plt.close()

plt.figure(figsize=(10,5))
sns.barplot(y='name', x='count', data=genre_stats.head(10), palette='magma')
plt.title('Топ-10 жанров')
plt.xlabel('Количество фильмов')
plt.ylabel('Жанр')
plt.tight_layout()
plt.savefig('data/top_genres.png')
plt.close()

plt.figure(figsize=(8,5))
sns.barplot(x='role', y='count', data=role_distribution, palette='coolwarm')
plt.title('Распределение ролей')
plt.xlabel('Роль')
plt.ylabel('Количество')
plt.tight_layout()
plt.savefig('data/role_distribution.png')
plt.close()

plt.figure(figsize=(10,5))
sns.barplot(y='full_name', x='count', data=actor_top10, palette='Blues_r')
plt.title('Топ-10 актёров')
plt.xlabel('Количество фильмов')
plt.ylabel('Актёр')
plt.tight_layout()
plt.savefig('data/actor_top10.png')
plt.close()


plt.figure(figsize=(10,5))
sns.histplot(rating_histogram, bins=20, kde=True, color='purple')
plt.title('Гистограмма рейтингов')
plt.xlabel('Рейтинг')
plt.ylabel('Количество')
plt.tight_layout()
plt.savefig('data/rating_histogram.png')
plt.close()

plt.figure(figsize=(8,5))
sns.barplot(x='type', y='count', data=type_distribution, palette='Set2')
plt.title('Распределение по типу')
plt.xlabel('Тип')
plt.ylabel('Количество')
plt.tight_layout()
plt.savefig('data/type_distribution.png')
plt.close()

print("Успешно завершено!")