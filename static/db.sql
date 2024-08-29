-- name: create_schema#
-- create the servers table
CREATE TABLE IF NOT EXISTS `servers` (
  "id" INTEGER PRIMARY KEY AUTOINCREMENT,
  "name" TEXT,
  "data_class" TEXT,
  "path" TEXT UNIQUE,
  "last_selected" BOOLEAN,
  "createAt" TEXT NOT NULL DEFAULT (datetime('now','localtime'))
);

-- name: insert!
-- 1. insert a new server into the database
INSERT INTO servers (name, data_class, path, last_selected) VALUES (:name, :data_class, :path, :last_selected);

-- name: delete_by_path!
-- 2. 根据路径删除服务器
DELETE FROM servers WHERE path = :path;

-- name: update_last_selected_by_path!
-- 3. 根据路径更新服务器的 last_selected
UPDATE servers SET last_selected = 1 WHERE path = :path;

-- name: update_last_selected!
-- 4. 修改 servers 表中 last_selected = 1 记录的状态
UPDATE servers
SET last_selected = 0
WHERE last_selected = 1;

-- name: select_names
-- 5. 选择去重后所有服务器的名称
SELECT DISTINCT name from servers;

-- name: select_data_classes
-- 6. 选择去重后所有服务器的数据类型
SELECT DISTINCT data_class from servers WHERE name = :name;

-- name: select_paths
-- 7. 选择服务器的路径
SELECT path from servers WHERE name = :name AND data_class = :data_class;

-- name: select_by_name^
-- 8. 根据服务器名称选择第一条记录
SELECT name, data_class, path, last_selected  from servers WHERE name = :name ORDER BY createAt;

-- name: select_id_by_path^
-- 9. 根据服务器路径选择服务器的id
SELECT id from servers WHERE path = :path;

-- name: select_by_last_selected^
-- 10. select all servers by the last selected server's path and name
SELECT name, data_class, path, last_selected FROM servers WHERE last_selected = 1;