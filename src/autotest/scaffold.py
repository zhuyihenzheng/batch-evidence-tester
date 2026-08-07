"""テストケースのひな形生成と複製。

ケースを 1 件増やすたびに YAML を手書きするのは手間だし、書き漏らしも起きる。
templates/ のひな形から生成するか、既存ケースを丸ごと複製して差分だけ直す。

設計方針:
  - YAML は「テキストとして」読み書きする。yaml.safe_load → yaml.dump を通すと
    コメントが全部消える。このリポジトリの YAML はコメントが説明書そのものなので、
    それを失うと生成物の価値がほとんど無くなる。
  - 生成物は必ず読み直して自己検証する。「生成できたが読めない」ものを黙って残さない。
  - 既存ファイルは上書きしない。取り違えで作業中のケースを消さないため。
  - 生成直後は enabled: false。作りかけのケースが全体の validate / run を
    巻き添えで落とさないようにする（準備ができたら人が true にする）。
"""

import shutil
from pathlib import Path
from typing import List, Optional

from .config import ConfigError, _load_yaml, find_case_files, validate_case_id

# Excel のシート名は 31 文字まで。超えると excel.py が切り詰めて連番を付けるため、
# 長い ID 同士が同じシート名に潰れて見分けが付かなくなる。
SHEET_NAME_LIMIT = 31


def templates_dir(project_root: Path) -> Path:
    return project_root / "templates"


def list_templates(project_root: Path) -> List[str]:
    """使えるひな形の名前一覧（拡張子なし）。"""
    directory = templates_dir(project_root)
    if not directory.is_dir():
        return []
    return sorted(p.stem for p in directory.glob("*.yaml"))


def _case_id_of(path: Path) -> str:
    """ケース定義ファイルの ID。id: 未指定ならファイル名（load_cases と同じ規則）。

    ここで stem へフォールバックしないと、id: を書いていないケースが
    「ID なし」に見えて衝突検査をすり抜ける。生成は成功するのに次の
    load_cases で ID 重複エラー、という分かりにくい失敗になる。
    """
    try:
        data = _load_yaml(path)
    except ConfigError:
        # 壊れた YAML は衝突検査の対象外。書式の誤りは load_cases が
        # ファイル名付きで報告するので、ここで巻き添えにしない
        return ""
    return str(data.get("id") or path.stem)


def _known_case_ids(cases_dir: Path) -> List[str]:
    """定義済みのケース ID をすべて集める。

    load_cases は enabled や絞り込みで件数が減るため、衝突検査には使えない。
    無効化されたケースと ID がぶつかっても証跡フォルダは衝突するので、
    ここでは「ファイルに書かれている ID すべて」を見る。
    """
    if not cases_dir.is_dir():
        return []
    return [cid for cid in (_case_id_of(p) for p in find_case_files(cases_dir)) if cid]


def _check_new_id(cases_dir: Path, case_id: str) -> None:
    """新しいケース ID として使えるかを検査する。使えなければ ConfigError。"""
    problems = validate_case_id(case_id)
    if problems:
        raise ConfigError("ケース ID が不正です:\n  - " + "\n  - ".join(problems))

    existing = _known_case_ids(cases_dir)
    if case_id in existing:
        raise ConfigError(
            "ケース ID '%s' は既に使われています。\n"
            "  証跡フォルダと Excel シートが上書きし合うため、別の ID にしてください。" % case_id)

    if len(case_id) > SHEET_NAME_LIMIT:
        raise ConfigError(
            "ケース ID が長すぎます（%d 文字 / 上限 %d 文字）: %s\n"
            "  Excel のシート名に使うため、これを超えると切り詰められて"
            "他のケースと見分けが付かなくなります。" % (len(case_id), SHEET_NAME_LIMIT, case_id))


def _case_yaml_path(cases_dir: Path, case_id: str) -> Path:
    return cases_dir / (case_id + ".yaml")


def _assert_absent(paths: List[Path]) -> None:
    """作成予定の場所が空いていることを確認する。1 つでも埋まっていたら中止。

    途中まで作ってから失敗すると、半端な生成物が残って次回の衝突検査に
    引っかかる。作る前に全部確かめる。
    """
    occupied = [p for p in paths if p.exists()]
    if occupied:
        raise ConfigError(
            "作成先が既に存在します（上書きはしません）:\n  - "
            + "\n  - ".join(str(p) for p in occupied))


def _verify_loadable(yaml_path: Path, case_id: str) -> None:
    """生成したケース定義がそのまま読めるかを確認する。

    読めないひな形を配って「生成はできたのに validate で落ちる」状態にしない。
    load_cases は使わない —— ひな形は準備が済むまで enabled: false で作るため、
    load_cases の絞り込みで消えてしまう。ファイル単体に同じ検証をかける。
    """
    from .config import _validate_case_schema

    try:
        data = _load_yaml(yaml_path)
    except ConfigError as exc:
        raise ConfigError(
            "生成したケースを読み込めませんでした。生成物は残してあるので確認してください:\n  %s"
            % str(exc).replace("\n", "\n  "))

    found = str(data.get("id") or yaml_path.stem)
    problems = _validate_case_schema(data, yaml_path, found)
    if problems:
        raise ConfigError(
            "生成したケース定義に問題があります: %s\n  - %s"
            % (yaml_path, "\n  - ".join(problems)))
    if case_id != found:
        raise ConfigError(
            "生成した %s の id が定義と一致しません（定義側: %s）。" % (case_id, found))


def create_case(project_root: Path, cases_dir: Path, case_id: str,
                template: str = "normal") -> List[Path]:
    """ひな形から新しいケースを生成する。作成したパスの一覧を返す。"""
    _check_new_id(cases_dir, case_id)

    available = list_templates(project_root)
    if template not in available:
        raise ConfigError(
            "ひな形 '%s' がありません。使えるひな形: %s\n  （%s）"
            % (template, available or "（なし）", templates_dir(project_root)))

    yaml_path = _case_yaml_path(cases_dir, case_id)
    input_dir = cases_dir / case_id / "input"
    expected_dir = project_root / "expected" / case_id
    _assert_absent([yaml_path, cases_dir / case_id, expected_dir])

    text = (templates_dir(project_root) / (template + ".yaml")).read_text(encoding="utf-8")
    # テンプレート中のプレースホルダはこの 1 種類だけ。str.format は使わない
    # （YAML 内の {date} や {dir: ...} まで巻き込んで壊れるため）
    text = text.replace("{case_id}", case_id)

    cases_dir.mkdir(parents=True, exist_ok=True)
    yaml_path.write_text(text, encoding="utf-8")
    input_dir.mkdir(parents=True, exist_ok=True)
    expected_dir.mkdir(parents=True, exist_ok=True)

    _verify_loadable(yaml_path, case_id)
    return [yaml_path, input_dir, expected_dir]


def _find_source_case(cases_dir: Path, src_id: str) -> Path:
    """複製元のケース定義ファイルを探す。"""
    for path in find_case_files(cases_dir):
        if _case_id_of(path) == src_id:
            return path
    raise ConfigError(
        "複製元のケース '%s' が見つかりません。定義済み: %s"
        % (src_id, sorted(_known_case_ids(cases_dir)) or "（なし）"))


def _copy_tree_if_exists(src: Path, dest: Path) -> Optional[Path]:
    if not src.is_dir():
        return None
    shutil.copytree(str(src), str(dest))
    return dest


def copy_case(project_root: Path, cases_dir: Path, src_id: str, new_id: str) -> List[Path]:
    """既存ケースを複製する。作成したパスの一覧を返す。

    資材（投入ファイル）・期待値・offline 用 fixtures もまとめて複製する。
    期待値まで複製するのは、古い期待値が残っていても出るのは NG であって
    偽 OK ではないため。「期待値が無くて実行できない」より安全側に倒す。
    """
    _check_new_id(cases_dir, new_id)
    src_yaml = _find_source_case(cases_dir, src_id)

    new_yaml = _case_yaml_path(cases_dir, new_id)
    src_material = src_yaml.parent / src_id
    new_material = src_yaml.parent / new_id
    src_expected = project_root / "expected" / src_id
    new_expected = project_root / "expected" / new_id
    src_fixtures = project_root / "fixtures" / src_id
    new_fixtures = project_root / "fixtures" / new_id
    _assert_absent([new_yaml, new_material, new_expected, new_fixtures])

    # 定義中の ID をすべて置換する。expected/<src_id>/... のようなパス参照も
    # まとめて新しい ID 側へ向く
    text = src_yaml.read_text(encoding="utf-8").replace(src_id, new_id)
    new_yaml.write_text(text, encoding="utf-8")

    created = [new_yaml]
    for src, dest in ((src_material, new_material),
                      (src_expected, new_expected),
                      (src_fixtures, new_fixtures)):
        copied = _copy_tree_if_exists(src, dest)
        if copied is not None:
            created.append(copied)

    # 資材フォルダが無い場合でも投入先は用意しておく（すぐ置けるように）
    if new_material not in created:
        (new_material / "input").mkdir(parents=True, exist_ok=True)
        created.append(new_material / "input")

    _verify_loadable(new_yaml, new_id)
    return created


def next_steps(project_root: Path, cases_dir: Path, case_id: str,
               config_hint: str = "") -> List[str]:
    """生成後に案内する手順。非開発者が次に何をすればよいか分かるように出す。"""
    config_opt = (" --config %s" % config_hint) if config_hint else ""
    return [
        "1) 投入ファイルを置く       : %s" % (cases_dir / case_id / "input"),
        "2) 定義を編集する           : %s" % _case_yaml_path(cases_dir, case_id),
        "3) enabled: true にする     : 準備ができるまでは false のまま（実行対象外）",
        "4) 設定を確認する           : python -m autotest validate%s" % config_opt,
        "5) 実行する                 : python -m autotest run --case %s%s" % (case_id, config_opt),
    ]
