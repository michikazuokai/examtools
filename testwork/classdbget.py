import sqlite3
import os
import yaml
from pathlib import Path

def read_and_process_yaml(file_path):
    """
    YAMLファイルを読み込み、その内容を処理する関数
    """
    try:
        # 'r' (読み込みモード) でファイルを開く
        with open(file_path, 'r', encoding='utf-8') as file:
            # yaml.safe_load() でファイルの内容をPythonの辞書/リスト構造に変換
            yaml_data = yaml.safe_load(file)
        # 🚨 ここでyaml_data（辞書/リスト）を呼び出し元に返す
        return yaml_data
    except FileNotFoundError:
        print(f"❌ エラー: ファイル '{file_path}' が見つかりません。")
    except yaml.YAMLError as e:
        print(f"❌ エラー: YAMLの解析中にエラーが発生しました。\n詳細: {e}")
    except Exception as e:
        print(f"❌ 予期せぬエラーが発生しました。\n詳細: {e}")

def get_nenji_by_subno(db_path, sub_no):
    """subNoからnenjiを取得"""
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT nennji FROM class WHERE subNo = ?", (sub_no,))
        result = cursor.fetchone()
        return result[0] if result else None
    except sqlite3.Error as e:
        print(f"エラー: {e}")
        return None
    finally:
        conn.close()

def get_name_by_stdno(db_path):
    """subNoからnenjiを取得"""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    sql_query = "SELECT stdNo, nickname FROM student"
    cursor.execute(sql_query)
    # fetchall() で結果をタプルのリストとして全て取得
    sql_results = cursor.fetchall()
    #print(f"🗃️ SQL実行結果 (タプルのリスト):\n{sql_results}")
    # 3. 辞書内包表記で変換
    # ----------------------------------------------------
    # row[0] が stdNo (キー)、row[1] が nickname (値)
    student_dict = {row[0]: row[1] for row in sql_results}
    print(student_dict)
    return student_dict

# 使用例
db_path = "/Volumes/NBPlan/TTC/カルテ管理/2025/DB/classdb.db"
sub_no = input("subNoを入力: ")  # 外部入力

nenji = get_nenji_by_subno(db_path, sub_no)

curdir = Path(__file__).parent
file_path = curdir / 'studentVersion2.yaml'
dt=read_and_process_yaml(file_path)
sdic=get_name_by_stdno(db_path)
keys_view = dt[2025][nenji].keys()
for k in keys_view:
    print(k)
    for v in dt[2025][nenji][k]['students']:
        name=sdic[str(v)]
        print(f"name: {name} ")
