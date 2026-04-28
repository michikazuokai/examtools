# versioncontrol_yaml.py
from pathlib import Path
from datetime import datetime
import yaml


def _get_slideinfo_path_from_inputpath(inputpath: str) -> tuple[Path, str]:
    """
    inputpath:
      .../1020701.GITバージョン管理/16/試験問題.xlsx

    return:
      slideinfo_path:
        .../1020701.GITバージョン管理/slideinfo/slideinfo.yaml
      koma_no:
        16
    """
    input_path = Path(inputpath)
    exam_dir = input_path.parent          # .../16
    koma_no = exam_dir.name               # "16"
    subject_dir = exam_dir.parent         # .../1020701.GITバージョン管理
    slideinfo_path = subject_dir / "slideinfo" / "slideinfo.yaml"

    return slideinfo_path, koma_no


def _load_slideinfo(slideinfo_path: Path) -> dict:
    if not slideinfo_path.exists():
        raise FileNotFoundError(f"slideinfo.yaml が見つかりません: {slideinfo_path}")

    with open(slideinfo_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    return data or {}


def _save_slideinfo(slideinfo_path: Path, data: dict) -> None:
    with open(slideinfo_path, "w", encoding="utf-8") as f:
        yaml.safe_dump(
            data,
            f,
            allow_unicode=True,
            sort_keys=False
        )


def init_db():
    """
    SQLite版との互換用。
    YAML版ではDB初期化は不要。
    """
    return None


def get_latest_version(inputpath: str, sheetname: str, conn=None) -> int:
    """
    SQLite版と同じ名前。
    slideinfo.yaml の該当コマから最新versionを取得する。
    """
    slideinfo_path, koma_no = _get_slideinfo_path_from_inputpath(inputpath)
    slideinfo = _load_slideinfo(slideinfo_path)

    info = slideinfo.get(str(koma_no), {})
    exam = info.get("exam", {})

    if exam.get("inputpath") != str(inputpath):
        return 0

    if exam.get("sheetname") != str(sheetname):
        return 0

    return int(exam.get("version", 0) or 0)


def check_version(hash_value: str, inputpath: str = "", sheetname: str = ""):
    """
    SQLite版では hash から履歴検索していた。
    YAML最新版のみ方式では、現在のhashと一致する場合だけ返す。

    互換的に使うなら inputpath と sheetname を渡す。
    """
    if not inputpath:
        return None

    slideinfo_path, koma_no = _get_slideinfo_path_from_inputpath(inputpath)
    slideinfo = _load_slideinfo(slideinfo_path)

    info = slideinfo.get(str(koma_no), {})
    exam = info.get("exam", {})

    if exam.get("hash") != hash_value:
        return None

    return (
        exam.get("hash"),
        exam.get("version"),
        exam.get("inputpath"),
        exam.get("sheetname"),
        exam.get("createdatetime"),
        exam.get("note", "")
    )


def insert_new_version(hash: str, inputpath: str, sheetname: str, note: str = "") -> int:
    """
    強制的にversionを+1してYAMLへ保存する。
    """
    slideinfo_path, koma_no = _get_slideinfo_path_from_inputpath(inputpath)
    slideinfo = _load_slideinfo(slideinfo_path)

    key = str(koma_no)
    if key not in slideinfo:
        raise KeyError(f"slideinfo.yaml に {key} がありません。")

    info = slideinfo[key]
    exam = info.get("exam", {})

    current_version = int(exam.get("version", 0) or 0)
    new_version = current_version + 1
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    info["exam"] = {
        "input_file": Path(inputpath).name,
        "inputpath": str(inputpath),
        "sheetname": str(sheetname),
        "version": new_version,
        "hash": hash,
        "createdatetime": now,
        "note": note,
    }

    info["update_at"] = now

    _save_slideinfo(slideinfo_path, slideinfo)

    return new_version


def ensure_version_entry(hash: str, inputpath: str, sheetname: str, note: str = "") -> int:
    """
    SQLite版と同じ名前で使える関数。

    - まだ exam 情報がない → version 1
    - hash が同じ → version はそのまま
    - hash が違う → version + 1
    """
    slideinfo_path, koma_no = _get_slideinfo_path_from_inputpath(inputpath)
    slideinfo = _load_slideinfo(slideinfo_path)

    key = str(koma_no)
    if key not in slideinfo:
        raise KeyError(f"slideinfo.yaml に {key} がありません。")

    info = slideinfo[key]
    exam = info.get("exam", {})

    current_hash = exam.get("hash")
    current_version = int(exam.get("version", 0) or 0)

    if current_hash == hash:
        return current_version

    new_version = current_version + 1
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    info["exam"] = {
        "input_file": Path(inputpath).name,
        "inputpath": str(inputpath),
        "sheetname": str(sheetname),
        "version": new_version,
        "hash": hash,
        "createdatetime": now,
        "note": note,
    }

    info["update_at"] = now

    _save_slideinfo(slideinfo_path, slideinfo)

    return new_version


def is_latest_version_for_file_sheet(hash: str, inputpath: str, sheetname: str) -> bool:
    """
    指定ファイル・シートの最新hashが指定hashと一致するか確認する。
    """
    slideinfo_path, koma_no = _get_slideinfo_path_from_inputpath(inputpath)
    slideinfo = _load_slideinfo(slideinfo_path)

    info = slideinfo.get(str(koma_no), {})
    exam = info.get("exam", {})

    return (
        exam.get("hash") == hash
        and exam.get("inputpath") == str(inputpath)
        and exam.get("sheetname") == str(sheetname)
    )


if __name__ == "__main__":
    # テスト例
    inputpath = "/Volumes/NBPlan/TTC/授業資料/2026年度/1020701.GITバージョン管理/16/試験問題.xlsx"
    print(ensure_version_entry("testhash2", inputpath, "1020701"))