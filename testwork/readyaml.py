import yaml
from pathlib import Path

# 1. ファイル名

def read_and_process_yaml(file_path):
    """
    YAMLファイルを読み込み、その内容を処理する関数
    """
    try:
        # 'r' (読み込みモード) でファイルを開く
        with open(file_path, 'r', encoding='utf-8') as file:
            # yaml.safe_load() でファイルの内容をPythonの辞書/リスト構造に変換
            yaml_data = yaml.safe_load(file)
        
        print("✅ YAMLファイルの読み込みに成功しました。")
        print("-" * 30)

        # 読み込んだデータの確認（ここでは処理はせず、単純に返す）
        print(f"型: {type(yaml_data)}")

        # 🚨 ここでyaml_data（辞書/リスト）を呼び出し元に返す
        return yaml_data
    except FileNotFoundError:
        print(f"❌ エラー: ファイル '{file_path}' が見つかりません。")
    except yaml.YAMLError as e:
        print(f"❌ エラー: YAMLの解析中にエラーが発生しました。\n詳細: {e}")
    except Exception as e:
        print(f"❌ 予期せぬエラーが発生しました。\n詳細: {e}")


if __name__ == "__main__":
    curdir = Path(__file__).parent
    file_path = curdir / 'studentVersion2.yaml'
    dt=read_and_process_yaml(file_path)
    print(len(dt[2025][1]['A']['students']))
    print(len(dt[2025][1]['B']['students']))
    print(len(dt[2025][2]['A']['students']))
    print(len(dt[2025][2]['B']['students']))

