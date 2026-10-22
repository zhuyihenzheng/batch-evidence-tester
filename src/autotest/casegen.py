"""テストケース md → ケース YAML（＋投入ファイル・初期データ SQL・期待値）の生成。

AI（Claude Code / GitHub Copilot / Codex）や人が書いたテストケース md を、決まった
規則で実行可能なケース定義へ変換する。変換そのものに LLM は使わない:
同じ md からは常に同じ YAML ができ、トークンも消費せず、書式の検証も確実になる。
LLM の仕事は「式样書を読んでテスト観点・テストデータ・期待値を md に書く」ところまで。

設計方針:
  - md の書式は docs/ai/testcase-format.md が唯一の仕様。ここはそれを機械的に実装する。
  - 未知の項目・解釈できない行はすべてエラー（行番号付き）。黙って読み飛ばすと
    「md に書いた確認が YAML に無い」= 検証が減ったまま誰も気づかない、になる。
  - 1 つでもエラーがあれば何も書かない（全部か無しか）。直す箇所を一度に全部返す。
  - 生成物は必ず enabled: false。AI が作った検証内容は、人がレビューして true に
    するまで実行対象にしない。
  - 生成物の 1 行目に生成マーカー（内容ハッシュ・生成元 md）を書き、生成した
    データファイルのハッシュも記録する。手で直したもの（ハッシュ不一致）は
    --overwrite でも上書きしない。直した内容を黙って巻き戻さないため。
  - 値を推測で補わない。補完するのは「書かなければ必ず不整合になる」対応関係だけ
    （DB 期待値 → 同じ列のスナップショット、期待ファイル → 同じファイルの回収）。
  - YAML は文字列として組み立てる。PyYAML 3.12（Anaconda 5.2）の dump は
    sort_keys を指定できずキー順が崩れるため。組み立てた後で読み直し、
    意図したデータと一致することを必ず確認する。
"""

import codecs
import csv
import hashlib
import io
import json
import os
import re
import shlex
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

from .config import (CASE_TOP_KEYS, ConfigError, DuplicateKeyError, _as_bool, _validate_case_schema,
                     find_case_files, load_settings, load_yaml_text, validate_case_id)
from .db import NULL_TEXT

# Excel のシート名上限（scaffold と同じ理由）
SHEET_NAME_LIMIT = 31

KINDS = {
    # 種別: (tags, mode)
    "正常系": (["正常系"], None),
    "異常系": (["異常系"], None),
    "環境不備": (["異常系", "環境不備"], None),
    "手動": (["手動"], "manual"),
}
DEFAULT_CLEAN = ["input_dir", "processed_dir", "error_dir", "output_dir"]
FRONT_KEYS = ("仕様書", "出力先", "既定クリア")

# ラベル → 値の種類。ここに無いラベルは綴り間違いとしてエラーにする
LABELS = {
    "種別": "kind",
    "目的": "text1",
    "仕様": "text1",
    "タグ": "words",
    "前提": "texts",
    "手順": "texts",
    "備考": "texts",
    "クリア": "aliases_or_none",
    "削除": "aliases",
    "準備SQL": "sqls",
    "入力": "inputs",
    "前置バッチ": "prebatches",
    "DB": "tables",
    "バッチ": "name1",
    "引数": "args",
    "基準日": "date",
    "回収": "collects",
    "撮影": "aliases_or_none",
    "終了コード": "int",
    "ジョブ状態": "jobstatus",
    "ログ必須": "literals",
    "ログ禁止": "literals",
    "ファイル": "counts",
    "確認": "texts",
    "後処理SQL": "sqls",
}
# 1 行（インライン）だけで書く項目。箇条書きの子項目を持たない
INLINE_ONLY = {"kind", "text1", "name1", "args", "date", "int", "jobstatus"}
# 子項目の続き行（インデントした通常行）を許す項目
CONTINUABLE = {"text1", "texts"}

BLOCK_ROLES = ("input", "seed", "expect_db", "expect_file", "file")
BLOCK_ATTRS = {
    "input": {"input", "encoding", "newline", "rename", "corrupt"},
    "seed": {"seed", "key", "delete", "identity", "fixture"},
    "expect_db": {"expect_db", "key", "ignore", "fixture"},
    "expect_file": {"expect_file", "name", "encoding", "ignore_line"},
    "file": {"file", "encoding", "newline"},
}
REPEATABLE_ATTRS = {"ignore_line"}

_ALIAS = re.compile(r"[A-Za-z_][A-Za-z0-9_]*\Z")
_H2 = re.compile(r"##(?!#)")
_HN = re.compile(r"#{1,6}(?:\s|$)")
# 箇条書きの下に字下げして置かれたコードブロックも受け付ける（LLM がよく書く形）
_FENCE = re.compile(r"(?P<indent> *)(?P<fence>`{3,}|~{3,})(?P<info>[^`]*)$")
_BULLET = re.compile(r"(?P<indent> *)[-*+](?:\s+(?P<body>.*))?$")
# 子項目だけは番号付き（1. / 1)）も箇条書きとして扱う
_ORDERED = re.compile(r"(?P<indent> *)\d+[.)](?:\s+(?P<body>.*))?$")
_RULE = re.compile(r"(?:-\s*){3,}$|(?:\*\s*){3,}$|(?:_\s*){3,}$")
_FIELD = re.compile(r"(?P<label>[^:：`]+?)\s*[:：]\s*(?P<value>.*)$")
_ATTR = re.compile(r'([A-Za-z_][A-Za-z0-9_]*)=("(?:[^"\\]|\\.)*"|\S*)')
_TABLE = re.compile(r"(?:\[[^\]]+\]|[^\W\d][\w$#@]*|#[\w$#@]*)"
                    r"(?:\.(?:\[[^\]]+\]|[^\W\d][\w$#@]*)){0,2}\Z")
_DATE = re.compile(r"\d{8}\Z|\d{4}-\d{2}-\d{2}\Z|\d{4}/\d{2}/\d{2}\Z")
# 引用符なしで書ける文字列（英字・かな漢字で始まり、記号は . - _ だけ）。
# 判定が甘くても dump_yaml の読み戻し確認で必ず検出される
_PLAIN = re.compile(r"[^\W\d][\w.\-]*\Z")
_YAML11_WORDS = {"y", "n", "yes", "no", "on", "off", "true", "false", "null", "~"}
_SQL_FILE = re.compile(r"[\w./\\-]+\.sql\Z", re.IGNORECASE)
_ENABLED_LINE = re.compile(r"^enabled:.*$", re.MULTILINE)
MARKER_PREFIX = "# generated-by: autotest gen"
_MARKER = re.compile(r"# generated-by: autotest gen sha256=(?P<sha>[0-9a-f]{64}) md=(?P<md>.+)$")
_GEN_FILE = re.compile(r"^# generated-file: sha256=(?P<sha>[0-9a-f]{64}) path=(?P<path>.+)$",
                       re.MULTILINE)


class GenError(ConfigError):
    """md の書式・内容の問題。行番号付きの一覧を持つ。"""

    def __init__(self, source: str, problems: List[str],
                 records: Optional[List[Dict[str, Any]]] = None) -> None:
        self.source = source
        self.problems = problems
        # --json で返す形（line / case / message）。AI が行番号で直接直せるように分けて持つ
        self.records = records if records is not None else [
            {"line": 0, "case": "", "message": text} for text in problems]
        super().__init__("%s\n  " % source + "\n  ".join(problems))


class _Problems:
    """問題を行番号付きで集める。最後にまとめて GenError にする。"""

    def __init__(self) -> None:
        self.items = []  # type: List[Tuple[int, str, str]]

    def add(self, line: int, message: str, case_id: str = "") -> None:
        self.items.append((line, case_id, message))

    def __bool__(self) -> bool:
        return bool(self.items)

    def _sorted(self) -> List[Tuple[int, str, str]]:
        return sorted(self.items, key=lambda x: x[0])

    def render(self) -> List[str]:
        out = []
        for line, case_id, message in self._sorted():
            text = ("%s: " % case_id if case_id else "") + message
            out.append("L%d %s" % (line, text) if line else text)
        return out

    def records(self) -> List[Dict[str, Any]]:
        return [{"line": line, "case": case_id, "message": message} for line, case_id, message in self._sorted()]

    def error(self, source: str) -> GenError:
        return GenError(source, self.render(), self.records())


# =============================================================================
# md の構文解析
# =============================================================================


class MdItem:
    """箇条書きの子項目 1 つ（続き行を含む）。"""

    def __init__(self, text: str, line: int, indent: int = 2) -> None:
        self.text = text
        self.line = line
        self.indent = indent


class MdField:
    """「- ラベル: 値」1 つ。"""

    def __init__(self, label: str, value: str, line: int) -> None:
        self.label = label
        self.kind = LABELS[label]
        self.value = value
        self.line = line
        self.items = []  # type: List[MdItem]


class MdBlock:
    """ケース節の中のコードブロック。"""

    def __init__(self, lang: str, attrs: List[Tuple[str, str]], lines: List[str], line: int) -> None:
        self.lang = lang
        self.attrs = attrs
        self.lines = lines
        self.line = line

    def attr(self, key: str, default: Optional[str] = None) -> Optional[str]:
        for k, v in self.attrs:
            if k == key:
                return v
        return default

    def attr_list(self, key: str) -> List[str]:
        return [v for k, v in self.attrs if k == key]

    @property
    def role(self) -> Optional[str]:
        roles = [k for k, _v in self.attrs if k in BLOCK_ROLES]
        return roles[0] if len(roles) == 1 else None


class MdCase:
    def __init__(self, case_id: str, name: str, line: int) -> None:
        self.case_id = case_id
        self.name = name
        self.line = line
        self.fields = {}  # type: Dict[str, MdField]
        self.blocks = []  # type: List[MdBlock]
        self.yaml_blocks = []  # type: List[MdBlock]

    def value(self, label: str) -> Optional[MdField]:
        return self.fields.get(label)


class MdDocument:
    def __init__(self, front: Dict[str, Any], cases: List[MdCase]) -> None:
        self.front = front
        self.cases = cases


def _dedent_raw(raw: str, indent: int) -> str:
    """コードブロックの中身の行から、開きフェンスの字下げ分だけ先頭の空白を除く。

    構造（フェンス・字下げ・箇条書き）の判定はタブを展開した行で行うが、中身は
    元の文字のまま残す。TSV やタブを含む値が空白に化けると、投入データや
    期待値が黙って変わるため。タブが字下げの境目をまたぐ分は空白で補う。
    """
    if indent == 0:
        return raw
    col = pos = 0
    while pos < len(raw) and col < indent and raw[pos] in " \t":
        col = col + 1 if raw[pos] == " " else (col // 4 + 1) * 4
        pos += 1
    return " " * max(0, col - indent) + raw[pos:]


def _intro_block_problem(info: str) -> Optional[str]:
    """最初の ## より前のコードブロック。データや YAML の追記なら、どのケースにも入らず消える。"""
    lang, attrs, junk = _split_info(info)
    if attrs or any("=" in word for word in junk) or lang.lower() in ("yaml", "yml"):
        return ("最初のケース（## ）より前のコードブロックはどのケースにも入りません。使うケースの ## の下へ"
                "移してください（複数ケースで共有するデータはファイルにして「準備SQL」「入力」から参照します）")
    return None


def _yaml_error_line(exc: Exception, first_line: int) -> int:
    """YAML の読み込みエラーを md の行番号にする（first_line = YAML の 1 行目の md 上の行）。"""
    mark = getattr(exc, "problem_mark", None)
    return first_line + mark.line if mark is not None else first_line


def _yaml_error_text(exc: Exception, first_line: int) -> str:
    if isinstance(exc, DuplicateKeyError):
        return "キー %r が 2 回書かれています（L%d と L%d）。後の値で前の値が消えるため 1 つにまとめてください" % (
            exc.key, first_line + exc.first_mark.line, first_line + exc.problem_mark.line)
    return str(getattr(exc, "problem", None) or exc).replace("\n", " ")


def _indent_of(line: str) -> int:
    return len(line) - len(line.lstrip(" "))


def _split_info(info: str) -> Tuple[str, List[Tuple[str, str]], List[str]]:
    """コードブロックの info string を (言語, 属性, 解釈できない語) に分ける。"""
    lang = ""
    attrs = []  # type: List[Tuple[str, str]]
    junk = []  # type: List[str]
    pos = 0
    text = info.strip()
    while pos < len(text):
        if text[pos].isspace():
            pos += 1
            continue
        m = _ATTR.match(text, pos)
        if m:
            raw = m.group(2)
            if len(raw) >= 2 and raw.startswith('"') and raw.endswith('"'):
                raw = raw[1:-1].replace('\\"', '"')
            attrs.append((m.group(1), raw))
            pos = m.end()
            continue
        end = pos
        while end < len(text) and not text[end].isspace():
            end += 1
        word = text[pos:end]
        if not lang and not attrs:
            lang = word
        else:
            junk.append(word)
        pos = end
    return lang, attrs, junk


def parse_markdown(text: str, problems: _Problems) -> MdDocument:
    """md を MdDocument にする。書式の問題は problems に積み、読める範囲は読み進める。"""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    front = {}  # type: Dict[str, Any]
    i = 0
    if lines and lines[0].strip() == "---":
        end = None
        for j in range(1, len(lines)):
            if lines[j].strip() in ("---", "..."):
                end = j
                break
        if end is None:
            problems.add(1, "フロントマター（---）が閉じていません")
            return MdDocument({}, [])
        try:
            loaded = load_yaml_text("\n".join(lines[1:end])) or {}
        except yaml.YAMLError as exc:
            # フロントマターは 2 行目から始まる
            problems.add(_yaml_error_line(exc, 2), "フロントマターを YAML として読めません: %s"
                         % _yaml_error_text(exc, 2))
            loaded = {}
        if not isinstance(loaded, dict):
            problems.add(1, "フロントマターは「キー: 値」で書いてください")
            loaded = {}
        unknown = sorted(str(k) for k in loaded if k not in FRONT_KEYS)
        if unknown:
            problems.add(1, "フロントマターに未知の項目があります: %s（使える項目: %s）"
                         % (unknown, ", ".join(FRONT_KEYS)))
        front = loaded
        i = end + 1

    cases = []  # type: List[MdCase]
    current = None  # type: Optional[MdCase]
    field = None  # type: Optional[MdField]
    cont = None  # type: Optional[Any]   # 続き行を受け付ける対象（MdItem or MdField）
    fence = None  # type: Optional[Tuple[str, int, str, int, List[str]]]
    in_comment = False
    comment_line = 0

    while i < len(lines):
        raw = lines[i]
        lineno = i + 1
        line = raw.expandtabs(4)
        i += 1

        if fence is not None:
            marker, indent, info, start, body = fence
            stripped = line.strip()
            # 閉じ側の字下げは開き側に合わせる（箇条書きの下に置いたブロックも閉じられるように）
            if (_indent_of(line) <= indent + 3 and stripped.startswith(marker[0] * len(marker))
                    and set(stripped) == {marker[0]}):
                fence = None
                if current is not None:
                    _add_block(current, info, body, start, problems)
                else:
                    problem = _intro_block_problem(info)
                    if problem:
                        problems.add(start, problem)
                continue
            body.append(_dedent_raw(raw, indent))
            continue

        if in_comment:
            end = line.find("-->")
            if end >= 0:
                in_comment = False
                if line[end + 3:].strip():
                    # 閉じた後ろの内容は読まれない。項目を書いても黙って消えるので止める
                    problems.add(lineno, "コメント（-->）の後ろに続けて書いた内容は読み込まれません。"
                                 "次の行に書いてください: %s" % line[end + 3:].strip()[:40],
                                 current.case_id if current else "")
            continue

        m = _FENCE.match(line)
        if m:
            # 導入部のコードブロックも読み進める（中の ## をケースの見出しにしない）。中身は捨てる
            fence = (m.group("fence"), len(m.group("indent")), m.group("info"), lineno, [])
            cont = None
            continue

        stripped = line.strip()
        if stripped.startswith("<!--"):
            # コメントアウトしたケース（<!-- ## TC… -->）を見出しとして拾わないよう、導入部でも読み飛ばす
            end = stripped.find("-->", 4)
            if end < 0:
                in_comment = True
                comment_line = lineno
            elif stripped[end + 3:].strip():
                problems.add(lineno, "コメント（<!-- … -->）の後ろに続けて書いた内容は読み込まれません。"
                             "コメントは単独の行に書いてください: %s" % stripped[end + 3:].strip()[:40],
                             current.case_id if current else "")
            continue
        if _H2.match(stripped) and _indent_of(line) <= 3:
            heading = stripped[2:]
            if heading and not heading[0].isspace():
                problems.add(lineno, "見出しは「## ケースID ケース名」の形で、## の後に空白を入れてください")
                continue
            parts = heading.strip().rstrip("#").strip().split(None, 1)
            if not parts:
                problems.add(lineno, "## の後にケース ID がありません")
                current = MdCase("", "", lineno)
                continue
            case_id = parts[0]
            name = parts[1].strip() if len(parts) > 1 else ""
            current = MdCase(case_id, name, lineno)
            cases.append(current)
            field = None
            cont = None
            if not name:
                problems.add(lineno, "ケース名がありません（「## %s 正常系：…」の形で書いてください）"
                             % case_id, case_id)
            continue

        if current is None:
            continue  # 導入部（見出し・説明・観点一覧表など）は自由記述

        if not stripped or (_RULE.match(stripped) and _indent_of(line) <= 3):
            cont = None   # 空行と区切り線（---）は読み飛ばす
            continue
        if _HN.match(stripped) and _indent_of(line) <= 3:
            problems.add(lineno, "ケースの中に見出しは書けません（ケースは ## で始まる 1 節だけです）",
                         current.case_id)
            continue

        b = _BULLET.match(line)
        if b and _indent_of(line) < 2:
            body = (b.group("body") or "").strip()
            fm = _FIELD.match(body)
            if not fm:
                problems.add(lineno, "項目は「- ラベル: 値」の形で書いてください: %s" % body, current.case_id)
                field = None
                cont = None
                continue
            label = fm.group("label").strip()
            if label not in LABELS:
                problems.add(lineno, "未知の項目「%s」（使える項目: %s）"
                             % (label, ", ".join(LABELS)), current.case_id)
                field = None
                cont = None
                continue
            if label in current.fields:
                problems.add(lineno, "「%s」が 2 回書かれています（L%d）。1 つにまとめてください"
                             % (label, current.fields[label].line), current.case_id)
                field = None
                cont = None
                continue
            field = MdField(label, fm.group("value").strip(), lineno)
            current.fields[label] = field
            cont = field if field.kind in CONTINUABLE else None
            continue

        sub = b or _ORDERED.match(line)
        if sub:
            body = (sub.group("body") or "").strip()
            indent = _indent_of(line)
            if isinstance(cont, MdItem) and indent >= cont.indent + 2 and field is not None \
                    and field.kind in CONTINUABLE:
                # 確認観点の中の入れ子の箇条書きは、その観点の続きとして残す
                cont.text += "\n" + line.strip()
                continue
            if field is None:
                problems.add(lineno, "どの項目の子項目か分かりません（直前に「- ラベル:」が必要です）",
                             current.case_id)
                continue
            if field.kind in INLINE_ONLY:
                problems.add(lineno, "「%s」は 1 行で書く項目です（子項目は書けません）" % field.label,
                             current.case_id)
                continue
            item = MdItem(body, lineno, indent)
            field.items.append(item)
            cont = item if field.kind in CONTINUABLE else None
            continue

        if cont is not None and _indent_of(line) >= 2:
            if isinstance(cont, MdItem):
                cont.text += "\n" + stripped
            else:
                cont.value = (cont.value + "\n" + stripped) if cont.value else stripped
            continue

        problems.add(lineno, "解釈できない行です（項目は「- ラベル: 値」、補足は「- 備考:」に書いてください）: %s"
                     % stripped[:40], current.case_id)

    if fence is not None:
        problems.add(fence[3], "コードブロックが閉じていません（%s）" % fence[0],
                     current.case_id if current else "")
    if in_comment:
        problems.add(comment_line, "コメント（<!--）が閉じていません", current.case_id if current else "")
    return MdDocument(front, cases)


def _add_block(case: MdCase, info: str, body: List[str], line: int, problems: _Problems) -> None:
    lang, attrs, junk = _split_info(info)
    if junk:
        problems.add(line, "コードブロックの属性を解釈できません: %s（属性は 名前=値 で書きます）"
                     % " ".join(junk), case.case_id)
        return
    block = MdBlock(lang, attrs, body, line)
    if not attrs:
        if lang.lower() in ("yaml", "yml"):
            case.yaml_blocks.append(block)
            return
        problems.add(line, "用途の分からないコードブロックです。データなら input= / seed= / expect_db= / "
                     "expect_file= / file= を、YAML の追記なら ```yaml を指定してください", case.case_id)
        return
    roles = [k for k, _v in attrs if k in BLOCK_ROLES]
    if len(roles) != 1:
        problems.add(line, "コードブロックには %s のどれか 1 つを指定してください"
                     % " / ".join(r + "=" for r in BLOCK_ROLES), case.case_id)
        return
    role = roles[0]
    allowed = BLOCK_ATTRS[role]
    seen = set()
    for key, _value in attrs:
        if key not in allowed:
            problems.add(line, "%s= のブロックに %s= は指定できません（使える属性: %s）"
                         % (role, key, ", ".join(sorted(allowed))), case.case_id)
            return
        if key in seen and key not in REPEATABLE_ATTRS:
            problems.add(line, "属性 %s= が 2 回指定されています" % key, case.case_id)
            return
        seen.add(key)
    case.blocks.append(block)


# =============================================================================
# 値の字句解析（`...` / → / = / 数値 / {オプション}）
# =============================================================================


class _TokenError(Exception):
    pass


def _read_span(text: str, pos: int) -> Tuple[str, int]:
    """pos から始まる `...`（CommonMark のコードスパン）を読む。"""
    run_end = pos
    while run_end < len(text) and text[run_end] == "`":
        run_end += 1
    width = run_end - pos
    search = run_end
    while True:
        k = text.find("`" * width, search)
        if k == -1:
            raise _TokenError("` が閉じていません")
        m = k
        while m < len(text) and text[m] == "`":
            m += 1
        if m - k == width:
            content = text[run_end:k]
            if len(content) >= 2 and content[0] == " " and content[-1] == " " and content.strip():
                content = content[1:-1]
            return content, m
        search = m


def _tokenize(text: str) -> List[Tuple[str, Any]]:
    tokens = []  # type: List[Tuple[str, Any]]
    pos = 0
    n = len(text)
    while pos < n:
        ch = text[pos]
        if ch.isspace():
            pos += 1
        elif ch == "`":
            content, pos = _read_span(text, pos)
            tokens.append(("span", content))
        elif text.startswith("->", pos):
            tokens.append(("arrow", "->"))
            pos += 2
        elif ch == "→":
            tokens.append(("arrow", ch))
            pos += 1
        elif ch == "=":
            tokens.append(("eq", ch))
            pos += 1
        elif ch in ",、，":
            tokens.append(("sep", ch))
            pos += 1
        elif ch == "{":
            raw = text[pos:].strip()
            try:
                value = load_yaml_text(raw)
            except yaml.YAMLError as exc:
                raise _TokenError("オプション {...} を YAML として読めません: %s" % str(exc).splitlines()[0])
            if not isinstance(value, dict):
                raise _TokenError("オプションは {キー: 値} の形で書いてください")
            tokens.append(("map", value))
            pos = n
        else:
            end = pos
            while end < n and not text[end].isspace() and text[end] not in "`=→{,、，" \
                    and not text.startswith("->", end):
                end += 1
            word = text[pos:end]
            tokens.append(("int", int(word)) if re.match(r"-?\d+\Z", word) else ("word", word))
            pos = end
    return tokens


class _Stream:
    def __init__(self, tokens: List[Tuple[str, Any]]) -> None:
        self.tokens = list(tokens)

    def peek(self) -> Optional[str]:
        return self.tokens[0][0] if self.tokens else None

    def take(self, *kinds) -> Any:
        if not self.tokens or self.tokens[0][0] not in kinds:
            return None
        return self.tokens.pop(0)[1]

    def options(self) -> Dict[str, Any]:
        value = self.take("map")
        return value if value is not None else {}

    def rest(self) -> str:
        return " ".join(str(v) for _k, v in self.tokens if _k != "sep")

    def done(self) -> bool:
        while self.tokens and self.tokens[0][0] == "sep":
            self.tokens.pop(0)
        return not self.tokens


def _dir_pattern(value: str) -> Tuple[str, str]:
    """`alias/pattern` を (論理名, パターン) に分ける。"""
    if "/" not in value:
        raise _TokenError("「論理名/パターン」の形で書いてください（例: `output_dir/RESULT_*.csv`）: %s" % value)
    alias, pattern = value.split("/", 1)
    if not _ALIAS.match(alias):
        raise _TokenError("論理名が不正です: %s" % alias)
    if not pattern:
        raise _TokenError("パターンが空です: %s" % value)
    return alias, pattern


def _shlex(value: str) -> List[str]:
    if "\\" in value:
        # shlex はバックスラッシュをエスケープとして消費する。Windows パスが黙って
        # 壊れた引数で batch を起動しないよう、ここで止める
        raise _TokenError("引数に \\ は使えません。パス区切りは / を使ってください: %s" % value)
    try:
        return shlex.split(value)
    except ValueError as exc:
        raise _TokenError("引数を分解できません（引用符の閉じ忘れ？）: %s" % exc)


def _field_entries(field: MdField) -> List[Tuple[str, int]]:
    """インライン値か子項目のどちらか一方を (文字列, 行番号) の並びで返す。"""
    if field.value and field.items:
        raise _TokenError("値を「%s:」の行と子項目の両方に書いています。どちらか一方にしてください" % field.label)
    if field.items:
        return [(item.text, item.line) for item in field.items]
    return [(field.value, field.line)] if field.value else []


# =============================================================================
# ケース 1 件分の組み立て
# =============================================================================


class GenFile:
    """生成するデータファイル 1 つ。"""

    def __init__(self, rel: str, path: Path, data: bytes, line: int) -> None:
        self.rel = rel          # プロジェクトルートからの相対（/ 区切り）
        self.path = path
        self.data = data
        self.line = line


class CasePlan:
    def __init__(self, md_case: MdCase) -> None:
        self.md = md_case
        self.case_id = md_case.case_id
        self.data = {}  # type: Dict[str, Any]
        self.files = []  # type: List[GenFile]
        self.yaml_path = None  # type: Optional[Path]
        self.material_dir = None  # type: Optional[Path]
        self.yaml_text = ""
        self.todos = []  # type: List[str]
        self.warnings = []  # type: List[str]
        self.ensure_dirs = []  # type: List[Path]
        self.action = ""
        self.note = ""
        self.stale = []  # type: List[Path]


class GenContext:
    def __init__(self, project_root: Path, cases_dir: Path, md_path: Path,
                 dialect=None, dialect_error: str = "settings が指定されていません") -> None:
        # seed / expect_db の SQL の書き方（SqlServerDialect / PostgresDialect）。
        # settings を読めなかったときは None で、SQL を作るブロックだけがエラーになる
        self.dialect = dialect
        self.dialect_error = dialect_error
        # 大文字小文字を区別しない FS（Windows / macOS）では、指定の綴りが違うだけで
        # 生成元 md の記録が変わり「別の md から生成された」と誤判定するため、実際の綴りに揃える
        self.project_root = _true_case(project_root)
        self.cases_dir = _true_case(cases_dir)
        self.md_path = _true_case(md_path)
        self.md_rel = rel_posix(self.md_path, self.project_root)

    def rel(self, path: Path) -> str:
        return rel_posix(path, self.project_root)

    def same_md(self, recorded: str) -> Tuple[bool, str]:
        """マーカーに記録された生成元がこの md か。(同じか, 補足) を返す。

        記録の md がもう存在しない場合は改名・移動とみなして引き継ぐ（内容ハッシュで
        手編集は別途守られている）。存在する別の md なら、同じ ID を 2 つの md が
        取り合っている本当の衝突なので False。
        """
        if recorded == self.md_rel:
            return True, ""
        candidate = Path(recorded)
        if not candidate.is_absolute():
            candidate = self.project_root / candidate
        if not candidate.exists():
            return True, "生成元の md が %s から変わっています（改名・移動とみなして引き継ぎます）" % recorded
        try:
            return os.path.samefile(str(candidate), str(self.md_path)), ""
        except OSError:
            return False, ""


def _true_case(path: Path) -> Path:
    """ディスク上の実際の綴り（大文字小文字）にしたパス。存在しない部分はそのまま残す。"""
    path = path.resolve()
    current = Path(path.parts[0])
    for part in path.parts[1:]:
        try:
            names = os.listdir(str(current))
        except OSError:
            names = []
        if part not in names:
            matches = [n for n in names if n.lower() == part.lower()]
            if len(matches) == 1:
                part = matches[0]
        current = current / part
    return current


def rel_posix(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return path.resolve().as_posix()


def _is_under(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _safe_relative(value: str) -> bool:
    if not value or "\\" in value or value.startswith("/") or re.match(r"[A-Za-z]:", value):
        return False
    parts = value.split("/")
    return all(part not in ("", ".", "..") for part in parts)


def _encode(lines: List[str], encoding: str, newline: str) -> bytes:
    text = "\n".join(lines) + ("\n" if lines else "")
    if newline == "crlf":
        text = text.replace("\n", "\r\n")
    return text.encode(encoding)


def _normalized_hash(data: bytes) -> str:
    """改行コードを LF に揃えたハッシュ。Windows の git（autocrlf）で CRLF に
    変換されただけのファイルを「手で編集された」と誤判定しないため。"""
    return hashlib.sha256(data.replace(b"\r\n", b"\n")).hexdigest()


def _parse_csv(block: MdBlock) -> Tuple[List[str], List[List[str]]]:
    """seed / expect_db の CSV。値の前後の空白は値の一部として扱う（何も削らない）。

    「1, 山田」のように , の後ろに空白を入れる書き方は LLM がよくするが、CSV では
    空白も値になる（' 山田' が投入・比較される）。黙って受け入れると初期データが
    意図と変わるので、引用符で囲まない空白はエラーにする。
    """
    text = "\n".join(block.lines)
    rows = [r for r in csv.reader(io.StringIO(text)) if r]
    if not rows:
        raise _TokenError("CSV にヘッダ行がありません")
    # skipinitialspace は「, の直後の引用符で囲まない空白」だけを読み飛ばす。結果が
    # 違う行に、囲まずに書いた空白がある（" 山田" と囲んだ空白は両方で同じになる）
    loose = [r for r in csv.reader(io.StringIO(text), skipinitialspace=True) if r]
    for number, (row, skipped) in enumerate(zip(rows, loose), start=1):
        if row != skipped:
            raise _TokenError("CSV の %d 行目で , の後ろに空白があります。空白も値の一部になります"
                              "（空白を含む値は \" \" で囲んでください）" % number)
    lines = [line for line in block.lines if line.strip()]
    if len(lines) == len(rows):   # 引用符の中の改行が無いとき、1 行 = 1 レコード
        for number, line in enumerate(lines, start=1):
            if line[:1] in (" ", "\t"):
                raise _TokenError("CSV の %d 行目が空白で始まっています。空白も先頭の値の一部になります"
                                  "（空白を含む値は \" \" で囲んでください）" % number)
    header = rows[0]
    if any(not h.strip() for h in header):
        raise _TokenError("CSV のヘッダに空の列名があります")
    if any(h != h.strip() for h in header):
        raise _TokenError("CSV のヘッダの列名に前後の空白があります: %s"
                          % [h for h in header if h != h.strip()])
    dupes = sorted({h for h in header if header.count(h) > 1})
    if dupes:
        raise _TokenError("CSV のヘッダに同じ列名があります: %s" % dupes)
    for offset, row in enumerate(rows[1:], start=2):
        if len(row) != len(header):
            raise _TokenError("CSV の %d 行目の列数が %d です（ヘッダは %d 列）" % (offset, len(row), len(header)))
    return header, rows[1:]


def _columns(value: Optional[str]) -> List[str]:
    return [c.strip() for c in (value or "").split(",") if c.strip()]


class SqlServerDialect(object):
    """SQL Server 用の SQL の書き方。

    値はすべて N'...' で渡し、型変換は SQL Server の暗黙変換に任せる（列型を
    知らなくても正しく入る）。
    """

    name = "sqlserver"
    label = "SQL Server"
    empty_effect = "数値列は 0、日付列は 1900-01-01 に変換されます"
    seed_note = ""
    terminator = ["GO"]

    @staticmethod
    def ident(name: str) -> str:
        return "[" + name.replace("]", "]]") + "]"

    @staticmethod
    def text(value: str) -> str:
        return "N'" + value.replace("'", "''") + "'"

    @staticmethod
    def check_table(table: str) -> Optional[str]:
        return None

    @staticmethod
    def insert_head(table: str, cols: str, identity: bool) -> str:
        return "INSERT INTO %s (%s) VALUES" % (table, cols)

    @staticmethod
    def before_insert(table: str, identity: bool) -> List[str]:
        return ["SET IDENTITY_INSERT %s ON;" % table] if identity else []

    @staticmethod
    def after_insert(table: str, header: List[str], identity: bool) -> List[str]:
        return ["SET IDENTITY_INSERT %s OFF;" % table] if identity else []


class PostgresDialect(object):
    """PostgreSQL 用の SQL の書き方。

    値は型を付けない '...' で渡し、PostgreSQL に投入先の列の型として解釈させる。
    N'...' は使えない —— PostgreSQL では character 型になり、varchar / text の列へ
    入れると末尾の空白が黙って落ちる（固定長・EDI のデータが変わる）。
    列名は "..." で囲む。大文字小文字がそのまま区別されるので、期待値 CSV の列名と
    スナップショットの列名が同じ綴りになる。
    """

    name = "postgresql"
    label = "PostgreSQL"
    empty_effect = "数値列・日付列ではエラーになります"
    seed_note = "PostgreSQL 用（settings の database.type）"
    terminator = []  # type: List[str]

    @staticmethod
    def ident(name: str) -> str:
        return '"' + name.replace('"', '""') + '"'

    @staticmethod
    def text(value: str) -> str:
        if "\\" in value:
            # standard_conforming_strings が off のサーバでも同じ値になるよう E'' で書く
            return "E'" + value.replace("\\", "\\\\").replace("'", "''") + "'"
        return "'" + value.replace("'", "''") + "'"

    @staticmethod
    def check_table(table: str) -> Optional[str]:
        if "[" in table or table.startswith("#"):
            return "PostgreSQL では表名を [ ] で囲んだり # で始めたりできません: %s" % table
        return None

    @staticmethod
    def insert_head(table: str, cols: str, identity: bool) -> str:
        # GENERATED ALWAYS の採番列は OVERRIDING SYSTEM VALUE が無いと値を入れられない
        return "INSERT INTO %s (%s)%s VALUES" % (table, cols, " OVERRIDING SYSTEM VALUE" if identity else "")

    @staticmethod
    def before_insert(table: str, identity: bool) -> List[str]:
        return []

    @classmethod
    def after_insert(cls, table: str, header: List[str], identity: bool) -> List[str]:
        # SQL Server は IDENTITY_INSERT で入れた値に合わせて採番を進めるが、PostgreSQL は
        # 明示した値を入れても採番（serial / identity のシーケンス）を進めない。そのままだと
        # batch の次の INSERT が投入済みの値と重複して失敗するので、投入した列のうち
        # 採番に紐づく列だけ、シーケンスを最大値より後ろへ進める（戻すことはしない）
        return [
            "-- 採番列（serial / identity）のシーケンスを投入した値より後ろへ進める",
            "DO $autotest$",
            "DECLARE",
            "  col text;",
            "  seq text;",
            "  top bigint;",
            "BEGIN",
            "  FOREACH col IN ARRAY ARRAY[%s] LOOP" % ", ".join(cls.text(c) for c in header),
            "    seq := pg_get_serial_sequence(%s, col);" % cls.text(table),
            "    IF seq IS NOT NULL THEN",
            "      EXECUTE format(%s, col) INTO top;" % cls.text("SELECT max(%%I) FROM %s" % table),
            "      IF top IS NOT NULL THEN",
            "        PERFORM setval(seq, GREATEST(top, nextval(seq)));",
            "      END IF;",
            "    END IF;",
            "  END LOOP;",
            "END",
            "$autotest$;",
        ]


DIALECTS = {d.name: d for d in (SqlServerDialect, PostgresDialect)}


def _sql_literal(value: str, dialect) -> str:
    if value == NULL_TEXT:
        return "NULL"
    return dialect.text(value)


def _chunks(items: List[Any], size: int) -> List[List[Any]]:
    return [items[i:i + size] for i in range(0, len(items), size)]


def _seed_sql(table: str, header: List[str], rows: List[List[str]], keys: List[str],
              delete: str, identity: bool, dialect) -> List[str]:
    """CSV の初期データを DELETE + INSERT の SQL にする（書き方は dialect）。

    (NULL) だけは NULL にする —— スナップショット側の NULL 表記と同じ記号にして、
    初期データと期待値を同じ書き方で書けるようにするため。
    """
    q = dialect.ident
    sql = ["-- 初期データ: %s（%d 行 / 削除: %s）" % (
        table, len(rows), {"key": "キー %s が一致する行" % ", ".join(keys),
                          "all": "全行", "none": "しない"}[delete])]
    if delete == "all":
        sql.append("DELETE FROM %s;" % table)
    elif delete == "key" and rows:
        positions = [header.index(k) for k in keys]
        if len(keys) == 1:
            for chunk in _chunks(rows, 1000):
                values = ", ".join(_sql_literal(r[positions[0]], dialect) for r in chunk)
                sql.append("DELETE FROM %s WHERE %s IN (%s);" % (table, q(keys[0]), values))
        else:
            for chunk in _chunks(rows, 200):
                conds = ["(" + " AND ".join("%s = %s" % (q(k), _sql_literal(r[p], dialect))
                                            for k, p in zip(keys, positions)) + ")" for r in chunk]
                sql.append("DELETE FROM %s WHERE %s;" % (table, " OR ".join(conds)))
    if rows:
        sql.extend(dialect.before_insert(table, identity))
        cols = ", ".join(q(c) for c in header)
        for chunk in _chunks(rows, 1000):   # SQL Server の VALUES 句は 1000 行が上限
            sql.append(dialect.insert_head(table, cols, identity))
            body = ["  (%s)" % ", ".join(_sql_literal(v, dialect) for v in row) for row in chunk]
            sql.append(",\n".join(body) + ";")
        sql.extend(dialect.after_insert(table, header, identity))
    sql.extend(dialect.terminator)
    return sql


def _dialect_for(ctx: "GenContext", table: str):
    """SQL を生成するブロック用。DB の種類で書き方が変わるので settings が読めないと作れない。"""
    if ctx.dialect is None:
        raise _TokenError("seed / expect_db の SQL は DB の種類（settings の database.type）で書き方が"
                          "変わります。settings を読み込めませんでした: %s" % ctx.dialect_error)
    problem = ctx.dialect.check_table(table)
    if problem:
        raise _TokenError(problem)
    return ctx.dialect


def _fixture_name(table: str) -> str:
    """offline の OfflineClient がフィクスチャ名に使う表名（_guess_table_name と同じ規則）。"""
    return table.strip("[]\"'").split(".")[-1]


def _merge_no_collision(base: Dict[str, Any], extra: Dict[str, Any], where: str,
                        problems: List[str]) -> None:
    """YAML ブロックの内容を足す。同じ項目を 2 か所で書いたらエラー（上書きも結合もしない）。

    以前の「defaults + cases」形式では、リストの丸ごと置換で継承した assert が黙って
    消えた（判定が減ったまま警告ゼロ = 偽 OK の方向）。ここでは衝突を必ず止める。
    """
    for key, value in extra.items():
        path = "%s.%s" % (where, key) if where else str(key)
        if key not in base:
            base[key] = value
        elif isinstance(base[key], dict) and isinstance(value, dict):
            _merge_no_collision(base[key], value, path, problems)
        else:
            problems.append("%s が md の項目と YAML ブロックの両方で指定されています。どちらか一方にまとめてください"
                            % path)


def build_case(md_case: MdCase, doc: MdDocument, ctx: GenContext, problems: _Problems) -> Optional[CasePlan]:
    """md のケース 1 件から YAML のデータと生成ファイルを組み立てる。"""
    plan = CasePlan(md_case)
    cid = md_case.case_id

    def err(line: int, message: str) -> None:
        problems.add(line, message, cid)

    # ID に問題があっても組み立ては続ける。他の項目の問題も一度に返すため
    for message in validate_case_id(cid):
        err(md_case.line, message)
    if len(cid) > SHEET_NAME_LIMIT:
        err(md_case.line, "ケース ID が長すぎます（%d 文字 / 上限 %d 文字。Excel のシート名に使うため）"
            % (len(cid), SHEET_NAME_LIMIT))

    group = str(doc.front.get("出力先") or "").strip().strip("/")
    group_dir = ctx.cases_dir / group if group and group != "." else ctx.cases_dir
    plan.yaml_path = group_dir / (cid + ".yaml")
    plan.material_dir = group_dir / cid

    f = md_case.fields
    for label in ("種別", "目的"):
        if label not in f or not f[label].value.strip():
            err(md_case.line, "「%s」は必須です" % label)

    kind_text = f["種別"].value.strip() if "種別" in f else ""
    if kind_text and kind_text not in KINDS:
        err(f["種別"].line, "種別は %s のいずれかです（指定値: %s）" % (" / ".join(KINDS), kind_text))
    kind_tags, mode = KINDS.get(kind_text, ([], None))

    # --- 値の取り出し（ラベルごとに字句解析）-----------------------------------
    parsed = {}  # type: Dict[str, Any]
    for label, field in f.items():
        try:
            parsed[label] = _parse_field(field)
        except _TokenError as exc:
            err(field.line, "「%s」: %s" % (label, exc))

    description = f["目的"].value.strip() if "目的" in f else ""
    if "仕様" in parsed:
        description += "（仕様: %s）" % parsed["仕様"]

    tags = list(kind_tags)
    for tag in parsed.get("タグ", []):
        if tag not in tags:
            tags.append(tag)

    data = {"id": cid, "name": md_case.name, "description": description,
            "tags": tags, "enabled": False}   # type: Dict[str, Any]
    if mode:
        data["mode"] = mode

    # --- setup ----------------------------------------------------------------
    removing = parsed.get("削除", [])
    if "クリア" in parsed:
        clean = parsed["クリア"]
    elif "既定クリア" in doc.front:
        clean = doc.front.get("既定クリア") or []
    else:
        clean = list(DEFAULT_CLEAN)
    clean = [a for a in clean if a not in removing]
    setup = {}  # type: Dict[str, Any]
    if clean:
        setup["clean_dirs"] = clean
    if removing:
        setup["remove_dirs"] = removing
    sql_items = list(parsed.get("準備SQL", []))
    input_items = list(parsed.get("入力", []))

    # --- データブロック ---------------------------------------------------------
    files = []  # type: List[GenFile]
    seeds = []  # type: List[str]
    expected_dbs = []  # type: List[Dict[str, Any]]
    auto_tables = []  # type: List[Dict[str, Any]]
    expected_files = []  # type: List[Dict[str, Any]]
    auto_collect = []  # type: List[Dict[str, Any]]
    db_labels = {t["name"] for t in parsed.get("DB", [])}

    def add_file(rel_root: Path, rel_inside: str, data_bytes: bytes, line: int) -> str:
        path = rel_root / rel_inside
        rel = ctx.rel(path)
        if any(g.rel == rel for g in files):
            err(line, "同じファイルを 2 回生成しようとしています: %s" % rel)
        files.append(GenFile(rel, path, data_bytes, line))
        return rel

    for block in md_case.blocks:
        try:
            _build_block(block, plan, ctx, add_file, seeds, expected_dbs, auto_tables,
                         expected_files, auto_collect, input_items, db_labels)
        except _TokenError as exc:
            err(block.line, "コードブロック: %s" % exc)

    if seeds:
        seed_path = plan.material_dir / "seed.sql"
        head = "-- 生成元: %s（%s）。md を直して gen し直すこと" % (ctx.md_rel, cid)
        if ctx.dialect is not None and ctx.dialect.seed_note:
            head += "。" + ctx.dialect.seed_note
        seed_text = "\n".join([head] + seeds)
        rel = add_file(plan.material_dir, "seed.sql", _encode(seed_text.split("\n"), "utf-8", "lf"),
                       md_case.line)
        if not _is_under(seed_path, ctx.project_root):
            # setup.sql の file はプロジェクトルート基準。外を指す絶対パスを YAML に残さない
            err(md_case.line, "ケースフォルダがプロジェクトの外にあるため seed.sql を相対参照できません")
        sql_items.append({"file": rel})

    if sql_items:
        setup["sql"] = sql_items
    if input_items:
        setup["input_files"] = input_items
    if "前置バッチ" in parsed:
        setup["batches"] = parsed["前置バッチ"]
    data["setup"] = setup

    tables = list(parsed.get("DB", [])) + auto_tables
    if tables:
        data["snapshot"] = {"tables": tables}

    execute = {}  # type: Dict[str, Any]
    for label, key in (("バッチ", "batch"), ("引数", "args"), ("基準日", "date")):
        if label in parsed:
            execute[key] = parsed[label]
    if execute:
        data["execute"] = execute

    collect = {}  # type: Dict[str, Any]
    collect_files = list(parsed.get("回収", []))
    for spec in auto_collect:
        if not any(c["dir"] == spec["dir"] and c["pattern"] == spec["pattern"] for c in collect_files):
            collect_files.append(spec)
    if collect_files:
        collect["files"] = collect_files
    if "撮影" in parsed:
        collect["folder_evidence"] = parsed["撮影"]
    if collect:
        data["collect"] = collect

    if "確認" in parsed:
        data["review"] = parsed["確認"]

    assertions = {}  # type: Dict[str, Any]
    if "終了コード" in parsed:
        if mode == "manual":
            err(f["終了コード"].line, "手動ケースでは終了コードを取得できないため「終了コード」は書けません"
                "（書くと確認したように見えて実際は SKIP になります）")
        assertions["exit_code"] = parsed["終了コード"]
    if "ジョブ状態" in parsed:
        assertions["job_status"] = parsed["ジョブ状態"]
    if expected_dbs:
        assertions["db"] = expected_dbs
    file_checks = list(parsed.get("ファイル", [])) + expected_files
    if file_checks:
        assertions["files"] = file_checks
    log = {}  # type: Dict[str, Any]
    if "ログ必須" in parsed:
        log["must_contain"] = parsed["ログ必須"]
    if "ログ禁止" in parsed:
        log["must_not_contain"] = parsed["ログ禁止"]
    if log:
        assertions["log"] = log
    if assertions:
        data["assert"] = assertions

    teardown_sql = parsed.get("後処理SQL")
    if teardown_sql:
        data["teardown"] = {"sql": teardown_sql}

    # --- YAML ブロック（md の項目で書けない設定の追記）---------------------------
    for block in md_case.yaml_blocks:
        try:
            extra = load_yaml_text("\n".join(block.lines)) or {}
        except yaml.YAMLError as exc:
            # ブロックの中身はフェンスの次の行から始まる。行番号は md の行に直して返す
            err(_yaml_error_line(exc, block.line + 1), "YAML ブロックを読めません: %s"
                % _yaml_error_text(exc, block.line + 1))
            continue
        if not isinstance(extra, dict):
            err(block.line, "YAML ブロックはマッピング（キー: 値）で書いてください")
            continue
        owned = [k for k in ("id", "name", "description", "tags", "enabled", "mode") if k in extra]
        if owned:
            err(block.line, "%s は md の見出し・種別・目的・タグから決まるため YAML ブロックには書けません"
                % ", ".join(owned))
            continue
        unknown = sorted(str(k) for k in extra if k not in CASE_TOP_KEYS)
        if unknown:
            err(block.line, "YAML ブロックに未知の項目があります: %s" % unknown)
            continue
        collisions = []  # type: List[str]
        _merge_no_collision(data, extra, "", collisions)
        for message in collisions:
            err(block.line, message)

    # --- 検証が 1 つも無いケースは作らない -------------------------------------
    has_check = bool(data.get("review")) or any(
        data.get("assert", {}).get(k) for k in ("exit_code", "job_status", "db", "files", "log")) \
        or "exit_code" in data.get("assert", {})
    if not has_check:
        err(md_case.line, "検証内容がありません（終了コード / ジョブ状態 / ログ必須 / ログ禁止 / ファイル / "
            "確認 / expect_db / expect_file のどれかを書いてください）。何も確認しないケースは"
            "実行しても合否の根拠になりません")

    # --- スキーマ検証（YAML を手で書いた場合と同じ規則）-------------------------
    for message in _validate_case_schema(data, plan.yaml_path, cid):
        err(md_case.line, "生成結果の検証: %s" % message)

    _collect_todos_and_warnings(plan, data, files, ctx)
    plan.data = data
    plan.files = files
    return plan


def _parse_field(field: MdField) -> Any:
    """ラベルの種類ごとに値を解釈する。問題があれば _TokenError。"""
    kind = field.kind
    if kind == "kind":
        return field.value.strip()
    if kind == "text1":
        if field.items:
            raise _TokenError("1 行（または続き行）で書いてください")
        return field.value.strip()
    if kind == "texts":
        entries = _field_entries(field)
        texts = [t.strip() for t, _l in entries if t.strip()]
        if not texts:
            raise _TokenError("内容がありません")
        return texts

    if kind in ("aliases", "aliases_or_none", "words", "literals", "sqls"):
        values = []  # type: List[Any]
        entries = _field_entries(field)
        if kind == "aliases_or_none" and len(entries) == 1 and entries[0][0].strip() == "なし":
            return []
        for text, _line in entries:
            for tkind, value in _tokenize(text):
                if tkind == "sep":
                    continue
                if kind == "literals":
                    if tkind != "span":
                        raise _TokenError("文言は `...` で囲んでください（%s）" % value)
                    if not value:
                        raise _TokenError("空の文言 `` は書けません")
                    values.append(value)
                elif kind == "sqls":
                    if tkind != "span":
                        raise _TokenError("SQL は `...` で囲んでください（ファイルなら `sql/xxx.sql`）")
                    values.append({"file": value} if _SQL_FILE.match(value) else value)
                elif tkind in ("span", "word") or (kind == "words" and tkind == "int"):
                    if kind != "words" and not _ALIAS.match(str(value)):
                        raise _TokenError("論理名が不正です: %s" % value)
                    values.append(str(value))
                else:
                    raise _TokenError("解釈できない値です: %s" % value)
        if not values and kind != "aliases_or_none":
            raise _TokenError("値がありません")
        return values

    if kind in ("name1", "date"):
        stream = _Stream(_tokenize(field.value))
        value = stream.take("span", "word", "int")
        if value is None or not stream.done():
            raise _TokenError("値を 1 つだけ書いてください")
        value = str(value)
        if kind == "date" and not _DATE.match(value):
            raise _TokenError("基準日は YYYYMMDD または YYYY-MM-DD で書いてください: %s" % value)
        return value
    if kind == "args":
        stream = _Stream(_tokenize(field.value))
        value = stream.take("span")
        if value is None or not stream.done():
            raise _TokenError("引数は 1 つの `...` にまとめて書いてください（例: `--mode daily --date {date}`）")
        return _shlex(value)
    if kind == "int":
        stream = _Stream(_tokenize(field.value))
        value = stream.take("int")
        if value is None or not stream.done():
            raise _TokenError("整数で書いてください")
        return value
    if kind == "jobstatus":
        stream = _Stream(_tokenize(field.value))
        status = stream.take("word", "span")
        job = stream.take("span", "word")
        if status is None or not stream.done():
            raise _TokenError("「COMPLETED」または「COMPLETED `ジョブ名`」の形で書いてください")
        return {"expected": status, "job": job} if job else status

    # --- 構造を持つ項目（1 行 = 1 件）-------------------------------------------
    # fixed は本体の書き方で決まる値（オプションで変えると本体と食い違うので禁止）、
    # defaults は省略時の値（{preview: false} のようにオプションで変えてよい）
    results = []  # type: List[Any]
    for text, _line in _field_entries(field):
        stream = _Stream(_tokenize(text))
        fixed = {}  # type: Dict[str, Any]
        defaults = {}  # type: Dict[str, Any]
        if kind == "inputs":
            src = stream.take("span")
            if src is None:
                raise _TokenError("投入元は `input/ファイル名` の形で書いてください")
            if not _safe_relative(src.replace("*", "x").replace("?", "x")):
                raise _TokenError("投入元はケースフォルダからの相対パスで書いてください: %s" % src)
            fixed["src"] = src
            if stream.take("arrow") is not None:
                dest = stream.take("span", "word")
                if dest is None or not _ALIAS.match(str(dest)):
                    raise _TokenError("→ の後に投入先の論理名を書いてください")
                fixed["dest_dir"] = str(dest)
            else:
                defaults["dest_dir"] = "input_dir"
        elif kind == "prebatches":
            name = stream.take("span", "word")
            if name is None:
                raise _TokenError("前置バッチの名前（settings の batches.<名前>）を書いてください")
            fixed["batch"] = str(name)
            args = stream.take("span")
            if args is not None:
                fixed["args"] = _shlex(args)
        elif kind == "tables":
            table = stream.take("span", "word")
            if table is None or not _TABLE.match(str(table)):
                raise _TokenError("テーブル名を `T_XXX` の形で書いてください")
            fixed["name"] = str(table)
            sql = stream.take("span")
            if sql is not None:
                fixed["sql"] = sql
            else:
                defaults["sql"] = "SELECT * FROM %s" % table
        elif kind == "collects":
            spec = stream.take("span")
            if spec is None:
                raise _TokenError("回収対象は `論理名/パターン` の形で書いてください")
            alias, pattern = _dir_pattern(spec)
            fixed.update({"dir": alias, "pattern": pattern})
            defaults["preview"] = True
        elif kind == "counts":
            spec = stream.take("span")
            if spec is None or stream.take("eq") is None:
                raise _TokenError("`論理名/パターン` = 件数 の形で書いてください（件数は必須）")
            count = stream.take("int")
            if count is None or count < 0:
                raise _TokenError("件数は 0 以上の整数で書いてください")
            alias, pattern = _dir_pattern(spec)
            defaults["name"] = "%s/%s が %d 件" % (alias, pattern, count)
            fixed["exists"] = {"dir": alias, "pattern": pattern, "count": count}
        else:  # pragma: no cover - LABELS と対応が取れていればここには来ない
            raise _TokenError("未対応の項目です")
        options = stream.options()
        if not stream.done():
            raise _TokenError("解釈できない値があります: %s" % stream.rest())
        clash = sorted(k for k in options if k in fixed)
        if clash:
            raise _TokenError("オプション %s は本体の書き方と重複しています" % ", ".join(clash))
        merged = dict(fixed)
        for key, value in defaults.items():
            merged[key] = options.pop(key, value)
        merged.update(options)
        # name は判定名として Excel の先頭に出るので、YAML でも先頭に置く
        item = {"name": merged.pop("name")} if "name" in merged and kind == "counts" else {}
        item.update(merged)
        results.append(item)
    if not results:
        raise _TokenError("値がありません")
    return results


def _build_block(block: MdBlock, plan: CasePlan, ctx: GenContext, add_file,
                 seeds: List[str], expected_dbs: List[Dict[str, Any]], auto_tables: List[Dict[str, Any]],
                 expected_files: List[Dict[str, Any]], auto_collect: List[Dict[str, Any]],
                 input_items: List[Dict[str, Any]], db_labels) -> None:
    role = block.role
    cid = plan.case_id
    encoding = block.attr("encoding", "utf-8")
    try:
        codecs.lookup(encoding)
    except LookupError:
        raise _TokenError("文字コード %s は使えません" % encoding)
    newline = block.attr("newline", "lf")
    if newline not in ("lf", "crlf"):
        raise _TokenError("newline は lf / crlf のどちらかです")

    def encoded(lines: List[str], enc: str = encoding, nl: str = newline) -> bytes:
        try:
            return _encode(lines, enc, nl)
        except UnicodeEncodeError as exc:
            bad = exc.object[exc.start:exc.end]
            raise _TokenError("%s で表せない文字があります: %r" % (enc, bad))

    expected_root = ctx.project_root / "expected" / cid
    fixtures_root = ctx.project_root / "fixtures" / cid

    if role == "input":
        alias, name = _dir_pattern(block.attr("input") or "")
        if "/" in name or not _safe_relative(name):
            raise _TokenError("input= のファイル名にフォルダは含められません（file= と「入力」を使ってください）")
        if any(ch in name for ch in "*?["):
            raise _TokenError("input= のファイル名に通配符は使えません")
        add_file(plan.material_dir, "input/" + name, encoded(block.lines), block.line)
        item = {"src": "input/" + name, "dest_dir": alias}
        if block.attr("rename"):
            item["rename"] = block.attr("rename")
        if block.attr("corrupt"):
            item["corrupt"] = block.attr("corrupt")
        if any(i.get("src") == item["src"] for i in input_items):
            raise _TokenError("同じ投入元 %s が「入力」にも書かれています" % item["src"])
        input_items.append(item)
        return

    if role == "file":
        target = block.attr("file") or ""
        if not _safe_relative(target):
            raise _TokenError("file= はケースからの相対パスで書いてください（.. や絶対パスは不可）: %s" % target)
        head, _sep, rest = target.partition("/")
        if head == "expected" and rest:
            add_file(expected_root, rest, encoded(block.lines), block.line)
        elif head == "fixtures" and rest:
            add_file(fixtures_root, rest, encoded(block.lines), block.line)
        else:
            add_file(plan.material_dir, target, encoded(block.lines), block.line)
        return

    if role == "seed":
        table = block.attr("seed") or ""
        if not _TABLE.match(table):
            raise _TokenError("seed= のテーブル名が不正です: %s" % table)
        dialect = _dialect_for(ctx, table)
        header, rows = _parse_csv(block)
        keys = _columns(block.attr("key"))
        delete = block.attr("delete") or ("key" if keys else "")
        if delete not in ("key", "all", "none"):
            raise _TokenError("seed には key=（そのキーの行を消して入れ直す）か delete=all / delete=none を"
                              "指定してください（既存データの消し方を明示するため）")
        if delete == "key":
            if not keys:
                raise _TokenError("delete=key には key= が必要です")
            missing = [k for k in keys if k not in header]
            if missing:
                raise _TokenError("key の列 %s が CSV のヘッダにありません" % missing)
            for row in rows:
                if any(row[header.index(k)] == NULL_TEXT for k in keys):
                    raise _TokenError("キー列に %s は書けません" % NULL_TEXT)
        if not rows and delete != "all":
            raise _TokenError("データ行がありません（テーブルを空にするだけなら delete=all）")
        identity = (block.attr("identity") or "off").lower()
        if identity not in ("on", "off"):
            raise _TokenError("identity は on / off のどちらかです")
        if any(v == "" for row in rows for v in row):
            plan.warnings.append(
                "L%d seed=%s に空欄があります。空文字として入ります（%s）。NULL にするなら %s と書いてください"
                % (block.line, table, dialect.empty_effect, NULL_TEXT))
        seeds.extend(_seed_sql(table, header, rows, keys, delete, identity == "on", dialect))
        fixture = block.attr("fixture")
        if fixture is not None:
            if fixture != "before":
                raise _TokenError("seed の fixture= は before だけです（offline 検証の実行前データ）")
            add_file(fixtures_root, "before_%s.csv" % _fixture_name(table),
                     encoded(block.lines, "utf-8", "lf"), block.line)
        return

    if role == "expect_db":
        table = block.attr("expect_db") or ""
        if not _TABLE.match(table):
            raise _TokenError("expect_db= のテーブル名が不正です: %s" % table)
        header, _rows = _parse_csv(block)
        keys = _columns(block.attr("key"))
        ignore = _columns(block.attr("ignore"))
        for label, cols in (("key", keys), ("ignore", ignore)):
            missing = [c for c in cols if c not in header]
            if missing:
                raise _TokenError("%s の列 %s が CSV のヘッダにありません" % (label, missing))
        if table not in db_labels:
            # 期待値の列だけを取る SELECT をこちらで組み立てるので、DB の種類が要る
            dialect = _dialect_for(ctx, table)
        if not keys and table not in db_labels:
            raise _TokenError("expect_db には key= が必要です（行の対応付けに使います。"
                              "並び順で比較したいときは「DB」に ORDER BY 付きの SQL を書いてください）")
        name = "db_%s.csv" % _fixture_name(table)
        rel = add_file(expected_root, name, encoded(block.lines, "utf-8", "lf"), block.line)
        item = {"table": table, "expected": rel}
        if keys:
            item["key"] = keys
        if ignore:
            item["ignore_columns"] = ignore
        expected_dbs.append(item)
        if table not in db_labels and not any(t["name"] == table for t in auto_tables):
            # 期待値の列だけを取る SELECT を補う。SELECT * だと期待値に無い列（個人情報を
            # 含み得る）まで証跡に出るうえ、並びが不定になる
            order = " ORDER BY %s" % ", ".join(dialect.ident(k) for k in keys) if keys else ""
            auto_tables.append({"name": table, "sql": "SELECT %s FROM %s%s"
                                % (", ".join(dialect.ident(c) for c in header), table, order)})
        fixture = block.attr("fixture")
        if fixture is not None:
            if fixture != "after":
                raise _TokenError("expect_db の fixture= は after だけです（offline 検証の実行後データ）")
            add_file(fixtures_root, "after_%s.csv" % _fixture_name(table),
                     encoded(block.lines, "utf-8", "lf"), block.line)
        return

    if role == "expect_file":
        alias, pattern = _dir_pattern(block.attr("expect_file") or "")
        name = block.attr("name")
        if not name:
            if any(ch in pattern for ch in "*?[{"):
                raise _TokenError("パターンに通配符・{date} があるときは name= で期待値ファイル名を指定してください")
            name = pattern.split("/")[-1]
        if "/" in name or not _safe_relative(name):
            raise _TokenError("name= はファイル名だけで書いてください: %s" % name)
        # 期待値ファイルは比較側が UTF-8 で読む。encoding= は実ファイル（batch の出力）の文字コード
        rel = add_file(expected_root, name, encoded(block.lines, "utf-8", "lf"), block.line)
        item = {"name": "%s/%s の内容" % (alias, pattern), "actual": {"dir": alias, "pattern": pattern},
                "expected": rel}
        if block.attr("encoding"):
            item["encoding"] = block.attr("encoding")
        patterns = block.attr_list("ignore_line")
        for regex in patterns:
            try:
                re.compile(regex)
            except re.error as exc:
                raise _TokenError("ignore_line の正規表現が不正です: %s (%s)" % (regex, exc))
        if patterns:
            item["ignore_line_patterns"] = patterns
        expected_files.append(item)
        auto_collect.append({"dir": alias, "pattern": pattern, "preview": True})
        return

    raise _TokenError("用途の分からないコードブロックです")  # pragma: no cover


def _collect_todos_and_warnings(plan: CasePlan, data: Dict[str, Any], files: List[GenFile],
                                ctx: GenContext) -> None:
    setup = data.get("setup", {})
    generated = {g.path.resolve() for g in files}
    for spec in setup.get("input_files", []):
        src = str(spec.get("src", ""))
        if any(ch in src for ch in "*?["):
            continue
        path = (plan.material_dir / src).resolve()
        if path not in generated and not path.exists():
            plan.todos.append("投入ファイルを置く: %s" % ctx.rel(path))
            plan.ensure_dirs.append(path.parent)
    for where, items in (("setup.sql", setup.get("sql", [])),
                         ("teardown.sql", data.get("teardown", {}).get("sql", []))):
        for spec in items:
            if isinstance(spec, dict) and spec.get("file"):
                path = (ctx.project_root / spec["file"]).resolve()
                if path not in generated and not path.exists():
                    plan.todos.append("%s の SQL を用意する: %s" % (where, spec["file"]))

    # 前ケースの残留物は「出力されたように見える」偽 OK の原因になる
    cleaned = {c if isinstance(c, str) else c.get("alias") for c in setup.get("clean_dirs") or []}
    cleaned |= set(setup.get("remove_dirs") or [])
    used = []
    for spec in setup.get("input_files", []):
        used.append((spec.get("dest_dir", "input_dir"), "投入先"))
    for spec in data.get("collect", {}).get("files", []):
        used.append((spec.get("dir"), "回収"))
    for item in data.get("assert", {}).get("files", []):
        loc = item.get("exists") or item.get("actual") or {}
        if item.get("exists") and loc.get("count") == 0:
            continue  # 0 件の確認は残留物があれば NG になる（偽 OK にはならない）
        used.append((loc.get("dir"), "ファイル判定"))
    reported = set()
    for alias, what in used:
        if alias and alias not in cleaned and alias not in reported:
            reported.add(alias)
            plan.warnings.append("%s（%s）が clean_dirs にありません。前のケースの残留ファイルで偽 OK に"
                                 "なる恐れがあります（「クリア」に加えるか、意図的なら無視）" % (alias, what))


# =============================================================================
# YAML の組み立て
# =============================================================================


def _is_scalar(value: Any) -> bool:
    return not isinstance(value, (dict, list))


def _scalar(value: Any) -> str:
    if value is True:
        return "true"
    if value is False:
        return "false"
    if value is None:
        return "null"
    if isinstance(value, int):
        return str(value)
    text = str(value)
    if _PLAIN.match(text) and text.lower() not in _YAML11_WORDS:
        return text
    # JSON の文字列は YAML の二重引用符スカラーとしてそのまま読める
    return json.dumps(text, ensure_ascii=False)


def _literal_ok(text: Any) -> bool:
    """複数行の文字列を | ブロックで出せるか（読み戻して同じになる形か）。"""
    if not isinstance(text, str) or "\n" not in text:
        return False
    lines = text.split("\n")
    return (not text.endswith("\n") and not text.startswith(" ") and "\t" not in text
            and "\r" not in text and all(line == line.rstrip() for line in lines))


def _width(text: str) -> int:
    """表示幅（全角 2 桁）。1 行に収まるかの判定に使う。"""
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(ch) in "FWA" else 1 for ch in text)


def _flow_list(values: List[Any]) -> Optional[str]:
    """短い値だけのリストを [a, b] 形式にする。1 つでも引用が要るなら全部引用して揃える。"""
    if not all(_is_scalar(v) and not _literal_ok(v) for v in values):
        return None
    if any(isinstance(v, str) and _width(v) > 40 for v in values):
        return None   # 確認観点のような文は 1 行 1 項目の方が読める
    rendered = [_scalar(v) for v in values]
    if any(r.startswith('"') for r in rendered):
        rendered = [json.dumps(v, ensure_ascii=False) if isinstance(v, str) else _scalar(v) for v in values]
    return "[%s]" % ", ".join(rendered)


def _emit_mapping(data: Dict[str, Any], indent: int) -> List[str]:
    pad = " " * indent
    out = []  # type: List[str]
    for key, value in data.items():
        k = _scalar(key)
        if isinstance(value, dict):
            flow = "{%s}" % ", ".join("%s: %s" % (_scalar(sk), _scalar(sv)) for sk, sv in value.items())
            if not value:
                out.append("%s%s: {}" % (pad, k))
            elif indent >= 2 and len(value) <= 4 and all(_is_scalar(v) and not _literal_ok(v)
                                                          for v in value.values()) \
                    and _width(pad + k + flow) <= 100:
                # exists / actual のような小さな対応は手書きのケースと同じく 1 行で出す
                out.append("%s%s: %s" % (pad, k, flow))
            else:
                out.append("%s%s:" % (pad, k))
                out.extend(_emit_mapping(value, indent + 2))
        elif isinstance(value, list):
            flow = _flow_list(value) if value else "[]"
            if flow is not None and _width(pad + k + flow) <= 100:
                out.append("%s%s: %s" % (pad, k, flow))
            else:
                out.append("%s%s:" % (pad, k))
                out.extend(_emit_sequence(value, indent + 2))
        elif _literal_ok(value):
            out.append("%s%s: |-" % (pad, k))
            out.extend(("%s  %s" % (pad, line)) if line else "" for line in value.split("\n"))
        else:
            out.append("%s%s: %s" % (pad, k, _scalar(value)))
    return out


def _emit_sequence(items: List[Any], indent: int) -> List[str]:
    pad = " " * indent
    out = []  # type: List[str]
    for item in items:
        if isinstance(item, dict) and item:
            sub = _emit_mapping(item, indent + 2)
            sub[0] = pad + "- " + sub[0][indent + 2:]
            out.extend(sub)
        elif isinstance(item, list):
            out.append(pad + "- " + json.dumps(item, ensure_ascii=False))
        elif _literal_ok(item):
            out.append(pad + "- |-")
            out.extend(("%s  %s" % (pad, line)) if line else "" for line in item.split("\n"))
        else:
            out.append(pad + "- " + _scalar(item))
    return out


def dump_yaml(data: Dict[str, Any]) -> str:
    """データを YAML 文字列にする。読み戻して一致しなければ例外（意図と違う定義を書かない）。"""
    text = "\n".join(_emit_mapping(data, 0)) + "\n"
    loaded = yaml.safe_load(text)
    if loaded != data:
        raise ConfigError("内部エラー: 生成した YAML を読み戻すと内容が変わります。\n%s" % text)
    return text


def _comment(text: str) -> List[str]:
    return [("# " + line).rstrip() for line in str(text).split("\n")]


def render_case(plan: CasePlan, ctx: GenContext) -> str:
    """生成マーカー・ヘッダコメント付きの YAML 全文を作る。"""
    md = plan.md
    head = ["# " + "=" * 77,
            "# %s %s" % (plan.case_id, md.name),
            "#   生成元 : %s:%d" % (ctx.md_rel, md.line)]
    for label in ("仕様", "前提", "手順", "備考"):
        field = md.fields.get(label)
        if field is None:
            continue
        if field.value:
            lines = field.value.split("\n")
            head.append(("#   %s : %s" % (label, lines[0])).rstrip())
            head.extend(("#          %s" % line).rstrip() for line in lines[1:])
        else:
            head.append("#   %s :" % label)
            for item in field.items:
                lines = item.text.split("\n")
                head.append(("#     - %s" % lines[0]).rstrip())
                head.extend(("#       %s" % line).rstrip() for line in lines[1:])
    head += [
        "#",
        "# この YAML は上の md から生成した。内容を変えるときは md を直して",
        "#   python -m autotest gen --md %s --overwrite" % ctx.md_rel,
        "# を実行する（YAML やデータファイルを直接直すと、以後 gen はこのケースを上書きしない）。",
        # 「enabled: false」をそのまま書かない。エディタの一括置換で true にされると、
        # この行まで変わって「手で編集された」と誤判定するため
        "# 生成直後の enabled は未レビューの印（false）。md と生成物を確認してから true にすること。",
        "# " + "=" * 77,
    ]
    for g in plan.files:
        head.append("# generated-file: sha256=%s path=%s" % (_normalized_hash(g.data), g.rel))
    body = "\n".join(head) + "\n" + dump_yaml(plan.data)
    return "%s sha256=%s md=%s\n%s" % (MARKER_PREFIX, _body_hash(body), ctx.md_rel, body)


def _body_hash(body: str) -> str:
    """生成マーカーに書く内容ハッシュ。enabled 行はレビューで変えてよい唯一の行なので除く。"""
    normalized = _ENABLED_LINE.sub("enabled: false", body.replace("\r\n", "\n"))
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _enabled_value(body: str) -> bool:
    """YAML 本文の enabled が真か。ケースの読み込み（config._as_bool）と同じ規則で読む。

    小文字の true だけを見ていたため、enabled: True / yes で有効にしたケースを --overwrite で
    false に戻しても「戻しました」と表示されず、ケースが黙って実行対象から消えていた。
    """
    m = re.search(r"^enabled:[ \t]*(?P<value>[^#\n]*?)[ \t]*(?:#.*)?$", body, re.MULTILINE)
    if m is None:
        return False
    try:
        return _as_bool(yaml.safe_load(m.group("value")) if m.group("value") else None, "enabled")
    except (ConfigError, yaml.YAMLError):
        return False


class _Existing:
    """既存 YAML の生成マーカーの読み取り結果。"""

    def __init__(self, text: str) -> None:
        text = text.replace("\r\n", "\n")
        first, _sep, body = text.partition("\n")
        m = _MARKER.match(first)
        self.generated = m is not None
        self.md = m.group("md").strip() if m else ""
        self.body_ok = bool(m) and _body_hash(body) == m.group("sha")
        self.body = body
        self.files = {fm.group("path").strip(): fm.group("sha") for fm in _GEN_FILE.finditer(body)} if m else {}
        self.enabled = _enabled_value(body)


# =============================================================================
# 計画と書き込み
# =============================================================================


def _existing_ids(cases_dir: Path) -> Dict[str, Path]:
    """定義済みケース ID（大文字小文字を区別しない）→ ファイル。Windows では同名扱いになるため。"""
    found = {}  # type: Dict[str, Path]
    if not cases_dir.is_dir():
        return found
    for path in find_case_files(cases_dir):
        try:
            loaded = yaml.safe_load(path.read_text(encoding="utf-8-sig")) or {}
        except (OSError, yaml.YAMLError, UnicodeDecodeError):
            loaded = {}
        cid = str(loaded.get("id") or path.stem) if isinstance(loaded, dict) else path.stem
        found[cid.lower()] = path
    return found


def _is_this_md(recorded: str, ctx: GenContext) -> bool:
    """記録の md がこの md そのものか（改名の引き継ぎはしない厳密な判定）。"""
    if recorded == ctx.md_rel:
        return True
    candidate = Path(recorded)
    if not candidate.is_absolute():
        candidate = ctx.project_root / candidate
    try:
        return candidate.exists() and os.path.samefile(str(candidate), str(ctx.md_path))
    except OSError:
        return False


def plan_generation(md_path: Path, ctx: GenContext, overwrite: bool) -> Tuple[List[CasePlan], List[str]]:
    """md を読み、ケースごとの生成計画を作る。問題があれば GenError（何も書かない）。"""
    problems = _Problems()
    try:
        text = md_path.read_bytes().decode("utf-8-sig")
    except UnicodeDecodeError:
        raise GenError(ctx.md_rel, ["UTF-8 として読めません（Shift_JIS で保存していませんか）"])
    doc = parse_markdown(text, problems)

    group = doc.front.get("出力先")
    if group is not None and (not isinstance(group, str) or (group.strip() not in ("", ".")
                                                             and not _safe_relative(group.strip().strip("/")))):
        problems.add(1, "出力先は cases フォルダからの相対パスで書いてください（.. や絶対パスは不可）: %r" % group)
        doc.front["出力先"] = ""
    default_clean = doc.front.get("既定クリア")
    if default_clean is not None and (not isinstance(default_clean, list) or any(
            not isinstance(a, str) or not _ALIAS.match(a) for a in default_clean)):
        problems.add(1, "既定クリア は論理名のリストで書いてください（例: [input_dir, output_dir]）")
        doc.front["既定クリア"] = list(DEFAULT_CLEAN)

    if not doc.cases and not problems:
        problems.add(0, "ケースがありません（「## ケースID ケース名」の見出しで 1 件以上書いてください）")

    seen = {}  # type: Dict[str, int]
    for case in doc.cases:
        key = case.case_id.lower()
        if key in seen:
            problems.add(case.line, "ケース ID が重複しています（L%d と同じ。大文字小文字の違いも同一扱い）"
                         % seen[key], case.case_id)
        seen.setdefault(key, case.line)

    plans = [build_case(case, doc, ctx, problems) for case in doc.cases if case.case_id]
    plans = [p for p in plans if p is not None]
    if problems:
        raise problems.error(ctx.md_rel)

    existing = _existing_ids(ctx.cases_dir)
    notices = []  # type: List[str]
    errors = _Problems()
    for plan in plans:
        try:
            plan.yaml_text = render_case(plan, ctx)
        except ConfigError as exc:
            errors.add(plan.md.line, str(exc), plan.case_id)
            continue
        other = existing.get(plan.case_id.lower())
        if other is not None and other.resolve() != plan.yaml_path.resolve():
            errors.add(plan.md.line, "ケース ID が既存のケースと重複しています: %s" % ctx.rel(other), plan.case_id)
            continue
        _decide_action(plan, ctx, overwrite, errors)

    # md から消えた生成済みケース（自動では消さない）
    ids = {p.case_id.lower() for p in plans}
    for cid, path in sorted(existing.items()):
        if cid in ids:
            continue
        try:
            head = path.read_text(encoding="utf-8-sig").split("\n", 1)[0]
        except (OSError, UnicodeDecodeError):
            continue
        m = _MARKER.match(head.rstrip("\r"))
        if m and _is_this_md(m.group("md").strip(), ctx):
            notices.append("md に無くなった生成済みケース（自動では削除しません）: %s" % ctx.rel(path))

    if errors:
        raise errors.error(ctx.md_rel)
    return plans, notices


def _add_note(plan: CasePlan, text: str) -> None:
    plan.note = (plan.note + " / " + text) if plan.note else text


def _decide_action(plan: CasePlan, ctx: GenContext, overwrite: bool, errors: _Problems) -> None:
    if not plan.yaml_path.exists():
        for g in plan.files:
            if g.path.exists() and _normalized_hash(g.path.read_bytes()) != _normalized_hash(g.data):
                errors.add(g.line, "生成先に内容の違う既存ファイルがあります（上書きしません）: %s" % g.rel,
                           plan.case_id)
        plan.action = "new"
        return

    existing = _Existing(plan.yaml_path.read_text(encoding="utf-8-sig"))
    if not existing.generated:
        plan.action = "manual"
        plan.note = "生成マーカーの無い（手で管理している）ケースのため更新しません"
        return
    same, moved_note = ctx.same_md(existing.md)
    if not same:
        errors.add(plan.md.line, "このケースは別の md（%s）から生成されています。同じ ID を 2 つの md に"
                   "書いていないか確認してください" % existing.md, plan.case_id)
        return
    plan.note = moved_note
    if not existing.body_ok:
        plan.action = "edited"
        _add_note(plan, "YAML が手で編集されているため更新しません（md の変更は反映されません）")
        return
    modified = []
    for rel, sha in sorted(existing.files.items()):
        path = ctx.project_root / rel
        if path.exists() and _normalized_hash(path.read_bytes()) != sha:
            modified.append(rel)
    if modified:
        plan.action = "edited"
        _add_note(plan, "生成したファイルが手で編集されているため更新しません: %s" % ", ".join(modified))
        return
    foreign = [g.rel for g in plan.files if g.rel not in existing.files and g.path.exists()
               and _normalized_hash(g.path.read_bytes()) != _normalized_hash(g.data)]
    if foreign:
        plan.action = "edited"
        _add_note(plan, "生成先に gen が作っていないファイルがあるため更新しません: %s" % ", ".join(foreign))
        return

    new_body = plan.yaml_text.split("\n", 1)[1]
    same_yaml = _body_hash(new_body) == _body_hash(existing.body)
    same_files = all(g.path.exists() and _normalized_hash(g.path.read_bytes()) == _normalized_hash(g.data)
                     for g in plan.files)
    stale = [ctx.project_root / rel for rel in existing.files if rel not in {g.rel for g in plan.files}]
    if same_yaml and same_files and not any(p.exists() for p in stale):
        plan.action = "unchanged"
        return
    if not overwrite:
        plan.action = "differs"
        _add_note(plan, "md と差分があります（--overwrite で更新）")
        return
    plan.action = "update"
    plan.stale = [p for p in stale if p.exists()]
    if existing.enabled:
        _add_note(plan, "enabled: false に戻しました（内容が変わったので再レビューしてください）")


_TMP_SUFFIX = ".autotest-gen-tmp"
_BACKUP_SUFFIX = ".autotest-gen-bak"


class GenWriteError(Exception):
    """書き込みの失敗。すべて元に戻したか（restored）と、戻せなかったファイルを持つ。"""

    def __init__(self, path: Path, cause: Exception, leftovers: List[Path]) -> None:
        self.path = path
        self.cause = cause
        self.leftovers = leftovers
        super().__init__("%s（%s）" % (path, cause))


def _sibling(path: Path, suffix: str) -> Path:
    return path.with_name(path.name + suffix)


def apply_plans(plans: List[CasePlan]) -> None:
    """計画どおりに書き込む。全ケース分を 1 つの単位として書き、途中で失敗したら元に戻す。

    以前は 1 ファイルずつ上書きしていたため、途中で失敗すると（Windows で期待値 CSV を
    Excel が開いている等）新旧のデータが混ざったまま残った。YAML に記録したハッシュとも
    合わなくなり、次の gen からは「手で編集された」と誤判定して更新もできなくなる。
    手順: ① 全内容を同じフォルダの一時ファイルへ書く → ② 既存ファイルを退避して差し替える
    （YAML は最後）→ ③ すべて成功したら退避を消す。①② のどこで失敗しても書く前の状態に戻す。
    """
    writes = []  # type: List[Tuple[Path, bytes]]
    removes = []  # type: List[Path]
    directories = []  # type: List[Path]
    for plan in plans:
        if plan.action not in ("new", "update"):
            continue
        writes.extend((g.path, g.data) for g in plan.files)
        removes.extend(plan.stale)
        directories.extend(plan.ensure_dirs)
    # YAML は最後に差し替える（失敗時に「無いファイルを参照する YAML」を一瞬でも作らないため）
    writes.extend((plan.yaml_path, plan.yaml_text.encode("utf-8"))
                  for plan in plans if plan.action in ("new", "update"))
    _commit_files(writes, removes)
    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)


def _commit_files(writes: List[Tuple[Path, bytes]], removes: List[Path]) -> None:
    staged = []  # type: List[Tuple[Path, Path]]
    replaced = []  # type: List[Tuple[Path, Optional[Path]]]   # (差し替えた先, 退避した元ファイル)
    removed = []  # type: List[Tuple[Path, Path]]
    current = None  # type: Optional[Path]
    try:
        for target, data in writes:
            current = target
            target.parent.mkdir(parents=True, exist_ok=True)
            tmp = _sibling(target, _TMP_SUFFIX)
            tmp.write_bytes(data)
            staged.append((target, tmp))
        for target, tmp in staged:
            current = target
            backup = None
            if target.exists():
                backup = _sibling(target, _BACKUP_SUFFIX)
                os.replace(str(target), str(backup))
            replaced.append((target, backup))
            os.replace(str(tmp), str(target))
        for path in removes:
            current = path
            backup = _sibling(path, _BACKUP_SUFFIX)
            os.replace(str(path), str(backup))
            removed.append((path, backup))
    except OSError as exc:
        leftovers = _rollback(replaced, removed)
        for _target, tmp in staged:
            _unlink_quietly(tmp)
        raise GenWriteError(current or Path("?"), exc, leftovers)
    for _target, backup in replaced:
        if backup is not None:
            _unlink_quietly(backup)
    for _path, backup in removed:
        _unlink_quietly(backup)


def _rollback(replaced: List[Tuple[Path, Optional[Path]]], removed: List[Tuple[Path, Path]]) -> List[Path]:
    """差し替え・削除を逆順に戻す。戻せなかった退避ファイルを返す（手で戻してもらう）。"""
    leftovers = []  # type: List[Path]
    for path, backup in reversed(removed):
        try:
            os.replace(str(backup), str(path))
        except OSError:
            leftovers.append(backup)
    for target, backup in reversed(replaced):
        try:
            if backup is not None:
                os.replace(str(backup), str(target))
            elif target.exists():
                target.unlink()
        except OSError:
            leftovers.append(backup or target)
    return leftovers


def _unlink_quietly(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


# =============================================================================
# CLI
# =============================================================================

_ACTION_LABEL = {
    "new": "新規", "update": "更新", "unchanged": "変更なし", "differs": "差分あり",
    "edited": "対象外", "manual": "対象外",
}


def _load_dialect(config_path: Optional[Path], project_root: Path):
    """settings の database.type から SQL の書き方を決める。(dialect, 読めなかった理由)"""
    if config_path is None:
        return None, "settings が指定されていません"
    try:
        settings = load_settings(config_path, project_root, require_runtime=False)
    except ConfigError as exc:
        return None, str(exc)
    return DIALECTS[settings.db_type], ""


def _print_json(data: Dict[str, Any]) -> None:
    # 1 行の JSON。AI の代理はこれだけ読めばよい（人向けの表を読み直さずに済む）
    print(json.dumps(data, ensure_ascii=False, separators=(",", ":")))


def main(args, project_root: Path, cases_dir: Path, config_path: Optional[Path] = None) -> int:
    """autotest gen --md <md> [--overwrite] [--dry-run] [--check] [--json] [--verbose]

    --check: 書き込まず、生成物が md と一致しているかだけを見る。一致しないケース
             （差分あり・手編集・未生成・md から消えたケース）があれば終了コード 1。
    --json : 結果を 1 行の JSON で出す（errors の line / case / message で直す場所が分かる）。
    """
    as_json = bool(getattr(args, "json", False))
    check = bool(getattr(args, "check", False))
    mode = "check" if check else ("dry-run" if args.dry_run else "write")

    def fail(message: str) -> int:
        if as_json:
            _print_json({"ok": False, "mode": mode, "errors": [{"line": 0, "case": "", "message": message}]})
        else:
            print("[設定エラー] %s" % message, file=sys.stderr)
        return 2

    if not args.md:
        return fail("--md でテストケース md を指定してください。"
                    "例: python -m autotest gen --md docs/testcases/order_import.md")
    if check and args.overwrite:
        return fail("--check は書き込まない確認用です。--overwrite と同時には指定できません")
    md_path = Path(args.md)
    if not md_path.is_file():
        return fail("テストケース md がありません: %s" % md_path)
    dialect, dialect_error = _load_dialect(config_path, project_root)
    ctx = GenContext(project_root, cases_dir, md_path, dialect, dialect_error)
    head = {"md": ctx.md_rel, "project": str(ctx.project_root), "db": dialect.name if dialect else None,
            "mode": mode}
    try:
        plans, notices = plan_generation(md_path, ctx, overwrite=args.overwrite and not check)
    except GenError as exc:
        if as_json:
            data = dict(head, ok=False, errors=exc.records, format="docs/ai/testcase-format.md")
            _print_json(data)
            return 2
        print("[形式エラー] %s（%d 件。ファイルは 1 つも書き込んでいません）" % (exc.source, len(exc.problems)),
              file=sys.stderr)
        for line in exc.problems:
            print("  %s" % line, file=sys.stderr)
        print("  書式: docs/ai/testcase-format.md", file=sys.stderr)
        return 2

    if not args.dry_run and not check:
        try:
            apply_plans(plans)
        except GenWriteError as exc:
            if as_json:
                _print_json(dict(head, ok=False, errors=[{"line": 0, "case": "", "message": "書き込みエラー: %s" % exc}],
                                 restored=not exc.leftovers, leftovers=[str(p) for p in exc.leftovers]))
                return 2
            print("[書き込みエラー] %s" % exc, file=sys.stderr)
            if exc.leftovers:
                print("  元に戻せなかったファイルがあります（%s の付いた退避から手で戻してください）:"
                      % _BACKUP_SUFFIX, file=sys.stderr)
                for path in exc.leftovers:
                    print("    %s" % path, file=sys.stderr)
            else:
                print("  書き込みを取り消し、すべて書く前の状態に戻しました。", file=sys.stderr)
            print("  Excel やエディタでそのファイルを開いていないか確認して、もう一度実行してください。",
                  file=sys.stderr)
            return 2

    active = [p for p in plans if p.action in ("new", "update", "unchanged", "differs")]
    warnings = ["%s: %s" % (p.case_id, w) for p in active for w in p.warnings]
    todos = ["%s: %s" % (p.case_id, t) for p in active for t in p.todos]
    out_of_sync = [p for p in plans if p.action != "unchanged"]
    in_sync = not out_of_sync and not notices
    # md の内容が生成物に反映されなかったケース。gen が成功を返して次の手順に並べると、
    # validate / run は古い YAML を検査・実行しているのに「md どおり」に見えてしまう
    blocked = [p for p in plans if p.action in ("differs", "edited", "manual")]

    targets = [p.case_id for p in plans if p.action in ("new", "update", "unchanged")]
    config_opt = (" --config %s" % args.config) if args.config else ""
    cases_opt = (" --cases-dir %s" % args.cases_dir) if args.cases_dir else ""
    case_opts = " ".join("--case %s" % c for c in targets)
    next_steps = []  # type: List[str]
    if targets and mode == "write":
        next_steps = [
            "python -m autotest validate --include-disabled %s%s%s" % (case_opts, config_opt, cases_opt),
            "md と生成物をレビューし、問題なければ各 YAML の enabled: false を true にする（AI は勝手に true にしない）",
            "python -m autotest run %s%s%s" % (case_opts, config_opt, cases_opt),
        ]

    if as_json:
        data = dict(head, ok=in_sync if check else not blocked, errors=[],
                    blocked=[p.case_id for p in blocked],
                    cases=[{"id": p.case_id, "action": p.action, "yaml": ctx.rel(p.yaml_path),
                            "files": len(p.files), "note": p.note} for p in plans],
                    warnings=warnings, todos=todos, notices=notices, next=next_steps)
        if check:
            data["in_sync"] = in_sync
        _print_json(data)
        if check:
            return 0 if in_sync else 1
        return 1 if blocked else 0

    def pad(text: str, width: int) -> str:
        return text + " " * max(1, width - _width(text))

    dry = "[dry-run] 書き込みはしていません。" if args.dry_run else ""
    dry = "[check] 書き込みはしていません。" if check else dry
    print("%sテストケース md: %s（%d ケース）" % (dry, ctx.md_rel, len(plans)))
    # 書き込み先を必ず見せる。--config の指定違いで別プロジェクトへ書いても気づけるように
    print("プロジェクト    : %s" % ctx.project_root)
    if dialect is not None and dialect.name != "sqlserver":
        # seed / 期待値の SQL の書き方が変わるので、既定以外のときは必ず見せる
        print("DB の種類       : %s" % dialect.label)
    id_w = max([len(p.case_id) for p in plans] + [8]) + 2
    for p in plans:
        label = _ACTION_LABEL.get(p.action, p.action) + ("（予定）" if args.dry_run and p.action in
                                                        ("new", "update") else "")
        extra = "  ファイル %d" % len(p.files) if p.files and p.action in ("new", "update") else ""
        print("  %s%s%s  %s%s" % (pad(label, 14), pad(p.case_id, id_w), ctx.rel(p.yaml_path), p.md.name, extra))
        if p.note:
            print("      ※ %s" % p.note)
        if args.verbose and p.action in ("new", "update", "differs"):
            for line in p.yaml_text.rstrip("\n").split("\n"):
                print("      | %s" % line)

    if warnings:
        print("警告:")
        for line in warnings:
            print("  %s" % line)
    if todos:
        print("残作業:")
        for line in todos:
            print("  %s" % line)
    for line in notices:
        print("※ %s" % line)

    if check:
        if in_sync:
            print("[同期チェック] OK: 生成物はすべて md と一致しています（%d ケース）" % len(plans))
            return 0
        print("[同期チェック] NG: md と一致していないものが %d 件あります"
              "（差分あり → gen --overwrite / 手編集・対象外 → md に反映してから作り直す）"
              % (len(out_of_sync) + len(notices)))
        return 1

    if blocked:
        print("[未反映] md の内容が生成物に反映されていないケースが %d 件あります: %s"
              % (len(blocked), ", ".join(p.case_id for p in blocked)))
        print("  差分あり → --overwrite で更新 / 対象外（手で編集・手で管理）→ 編集内容を md に移し、"
              "その YAML と生成ファイルを消してから gen")
        print("  これらは validate / run しても md の内容を検査したことにならないため、次の手順から外しています。")
    if next_steps:
        print("次の手順:")
        print("  1) 静的検査: %s" % next_steps[0])
        print("  2) %s" % next_steps[1])
        print("  3) 実行    : %s" % next_steps[2])
    return 1 if blocked else 0
