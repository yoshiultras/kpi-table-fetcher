import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()


class Database:

    # Все данные для отображения метрик в таблице
    # section_id для определения под какую секцую записывать метрику
    selectMetricsSQL = """SELECT metric_number, 
                       metric_subnumber, 
                       md.description, 
                       unit_of_measurement, 
                       base_level, 
                       average_level, 
                       goal_level, 
                       measurement_frequency, 
                       conditions, 
                       notes, 
                       points, 
                       section_id 
                       FROM metric_descriptions AS md 
                       JOIN sections AS s ON md.section_id = s.id 
                       ORDER BY 1, 2"""

    # Данные для формирования секций таблицы
    selectSectionsSQL = """SELECT id, description FROM sections;"""

    # Логика подключения к БД
    _connection = None

    @classmethod
    def get_connection(cls):
        if cls._connection is None:
            try:
                cls._connection = psycopg2.connect(
                    dbname=os.getenv("DB_NAME"),
                    user=os.getenv("DB_USER"),
                    password=os.getenv("DB_PASSWORD"),
                    host=os.getenv("DB_HOST"),
                    port=os.getenv("DB_PORT"),
                )
                print("Соединение с PostgreSQL установлено")
            except Exception as error:
                print("Ошибка при подключении к PostgreSQL", error)
                cls._connection = None
        return cls._connection

    @classmethod
    def close_connection(cls):
        if cls._connection:
            cls._connection.close()
            print("Соединение с PostgreSQL закрыто")
            cls._connection = None

    # Метод получения метрик
    @classmethod
    def get_metrics(cls):
        with Database.get_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(cls.selectMetricsSQL)
                results = cursor.fetchall()
                return results

    # Метод подключения групп
    @classmethod
    def get_sections(cls):
        with Database.get_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(cls.selectSectionsSQL)
                results = cursor.fetchall()
                return results
