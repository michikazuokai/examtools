import sqlite3
from pathlib import Path
from datetime import datetime

# データベースのパス
curdir = Path(__file__).parent.parent
db_path = curdir / 'db/versions.db'

# 1. DB初期化
def init_db():
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path, timeout=10) as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS versions (
                hash TEXT PRIMARY KEY,
                version INTEGER,
                inputpath TEXT,
                sheetname TEXT,
                createdatetime TEXT,
                note TEXT
            )
        ''')

# 2. ハッシュ存在確認
def check_version(hash_value):
    with sqlite3.connect(db_path, timeout=10) as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM versions WHERE hash = ?", (hash_value,))
        result = cur.fetchone()
    return result

# 3. 最新バージョン取得（接続を外部から渡せるように）
def get_latest_version(inputpath: str, sheetname: str, conn=None) -> int:
    should_close = False
    if conn is None:
        conn = sqlite3.connect(db_path, timeout=10)
        should_close = True

    try:
        cur = conn.cursor()
        cur.execute(
            '''
            SELECT MAX(version) FROM versions
            WHERE inputpath = ? AND sheetname = ?
            ''',
            (inputpath, sheetname)
        )
        row = cur.fetchone()
        return row[0] if row and row[0] is not None else 0
    except sqlite3.Error as e:
        print(f"❌ SQLiteエラーが発生しました: {e}")
        return 0
    finally:
        if should_close:
            conn.close()

# 4. 新しいバージョンを挿入（接続を共通化）
def insert_new_version(hash: str, inputpath: str, sheetname: str, note: str = "") -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with sqlite3.connect(db_path, timeout=10) as conn:
        current_version = get_latest_version(inputpath, sheetname, conn)
        new_version = current_version + 1

        cur = conn.cursor()
        cur.execute(
            '''
            INSERT INTO versions (hash, version, inputpath, sheetname, createdatetime, note)
            VALUES (?, ?, ?, ?, ?, ?)
            ''',
            (hash, new_version, inputpath, sheetname, now, note)
        )
        return new_version

# 5. エントリがなければ登録
def ensure_version_entry(hash: str, inputpath: str, sheetname: str, note: str = "") -> int:
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            cur = conn.cursor()
            cur.execute('SELECT version FROM versions WHERE hash = ?', (hash,))
            row = cur.fetchone()

            lstver = get_latest_version(inputpath, sheetname, conn)

            if not row or lstver > row[0]:
                # エントリがなければ追加
                current_version = lstver + 1
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cur.execute(
                    '''
                    INSERT INTO versions (hash, version, inputpath, sheetname, createdatetime, note)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ''',
                    (hash, current_version, inputpath, sheetname, now, note)
                )
                return current_version
            else:
                return row[0]
    except sqlite3.Error as e:
        print(f"❌ SQLiteエラー (ensure_version_entry): {e}")
        raise

# 6. 指定ファイル・シートの最新バージョンがこのハッシュか確認
def is_latest_version_for_file_sheet(hash, inputpath, sheetname) -> bool:
    try:
        with sqlite3.connect(db_path, timeout=10) as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT hash FROM versions
                WHERE inputpath = ? AND sheetname = ?
                ORDER BY version DESC
                LIMIT 1
            """, (inputpath, sheetname))

            row = cur.fetchone()
    except sqlite3.Error as e:
        print(f"❌ SQLiteエラー (is_latest_version_for_file_sheet): {e}")
        return False

    if row is None:
        return False

    latest_hash = row[0]
    return hash == latest_hash

# テスト用（直接実行時）
if __name__ == "__main__":
    print(ensure_version_entry('oohvore', 'aaa/xxxx', '484500'))