import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()


class Database:
    _connection = None

    @classmethod
    def get_connection(cls):
        if cls._connection is None:
            try:
                cls._connection = psycopg2.connect(
                    dbname=os.getenv('DB_NAME'),
                    user=os.getenv('DB_USER'),
                    password=os.getenv('DB_PASSWORD'),
                    host=os.getenv('DB_HOST'),
                    port=os.getenv('DB_PORT'),
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


# Пример использования
if __name__ == "__main__":
    connection = Database.get_connection()
    cursor = connection.cursor()

    # Выполнение SQL-запроса
    cursor.execute("SELECT version();")
    db_version = cursor.fetchone()
    print("Вы подключены к - ", db_version)

    cursor.close()
    Database.close_connection()