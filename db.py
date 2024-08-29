import sqlite3
import typing as tp
from contextlib import contextmanager
from dataclasses import dataclass

import aiosql

from utils import resource

sql = resource("./static/db.sql")
db = "./data.db"


def initial_db():
    queries = aiosql.from_path(sql, "sqlite3", encoding="utf-8")

    with sqlite3.connect(db) as conn:
        queries.create_schema(conn)

    return queries


queries = initial_db()


@dataclass
class Server:
    name: str
    data_class: str
    path: str
    last_selected: bool = False


class SqliteConn:
    def __init__(self, db: str) -> None:
        self.conn = sqlite3.connect(db)

    def execute(self, sql: str, params: tp.Tuple = ()) -> sqlite3.Cursor:
        cursor = self.conn.cursor()
        cursor.execute(sql, params)

        self.conn.commit()
        results = cursor.fetchall()
        cursor.close()
        return results

    def insert(self, server: Server):
        queries.insert(
            self.conn, server.name, server.data_class, server.path, server.last_selected
        )

    def delete_by_path(self, path: str) -> None:
        queries.delete_by_path(self.conn, path)

    def update_last_selected(self) -> None:
        queries.update_last_selected(self.conn)

    def update_last_selected_by_path(self, path) -> None:
        queries.update_last_selected_by_path(self.conn, path)

    def select_id_by_path(self, path: str) -> tp.Optional[int]:
        return queries.select_id_by_path(self.conn, path)

    def select_names(self) -> tp.List[str]:
        return [each[0] for each in queries.select_names(self.conn)]

    def select_data_classes_by_name(self, name: str) -> tp.List[str]:
        return [each[0] for each in queries.select_data_classes(self.conn, name)]

    def select_paths(self, name: str, data_class: str) -> tp.List[str]:
        return [each[0] for each in queries.select_paths(self.conn, name, data_class)]

    def select_by_name(self, name: str) -> Server:
        info = queries.select_by_name(self.conn, name)
        if not info:
            return None

        server = Server(*info)
        self.update_last_selected_by_path(server.path)
        return server

    def select_by_last_selected(self) -> Server:
        info = queries.select_by_last_selected(self.conn)
        if not info:
            return None
        return Server(*info)

    def close(self) -> None:
        self.conn.commit()
        self.conn.close()


@contextmanager
def connection():
    try:
        conn = SqliteConn(db)
        yield conn
    finally:
        # 模拟资源释放
        conn.close()
