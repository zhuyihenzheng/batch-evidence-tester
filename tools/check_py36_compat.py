# -*- coding: utf-8 -*-
"""ソースが Python 3.6 で動く構文だけで書かれているか検査する。

対象環境が Anaconda 5.2（Python 3.6.5）のため、3.7 以降の機能が
混入すると import 時点で落ちる。手元に 3.6 が無くても検出できるよう、
AST を走査して 3.7+ 固有の構文・API を洗い出す。

  python tools/check_py36_compat.py src/autotest demo
"""

import ast
import sys
import os

# PEP 585 の組み込みジェネリクス（3.9+）。3.6 では typing.List 等を使う必要がある
BUILTIN_GENERICS = {"list", "dict", "tuple", "set", "frozenset", "type"}

# 3.7 以降で追加されたキーワード引数
NEW_KWARGS = {
    "capture_output": "subprocess.run(capture_output=) は Python 3.7+",
    "text": "subprocess.run(text=) は Python 3.7+",
    "onexc": "shutil.rmtree(onexc=) は Python 3.12+（バージョン分岐済みなら可）",
}

# 新しめのライブラリ API
NEW_ATTRS = {
    "textlength": "ImageDraw.textlength は Pillow 8.0+",
    "textbbox": "ImageDraw.textbbox は Pillow 8.0+",
    "removeprefix": "str.removeprefix は Python 3.9+",
    "removesuffix": "str.removesuffix は Python 3.9+",
}


class Py36Checker(ast.NodeVisitor):
    def __init__(self, path, source):
        self.path = path
        self.lines = source.splitlines()
        self.problems = []
        self._guarded_lines = self._collect_guarded_lines(source)

    def _collect_guarded_lines(self, source):
        """実行時ガードで守られている行番号を集める。

        次の 2 つは 3.6 で実行されないため見逃してよい:
          - `if sys.version_info >= (3, x):`  … Python バージョン分岐
          - `if hasattr(obj, "new_api"):`     … 新しい API の存在確認
        いずれも「無ければ古い書き方にフォールバックする」意図の記述。
        """
        guarded = set()
        try:
            tree = ast.parse(source)
        except SyntaxError:
            return guarded
        for node in ast.walk(tree):
            if not isinstance(node, ast.If):
                continue
            test_dump = ast.dump(node.test)
            if "version_info" in test_dump or "'hasattr'" in test_dump:
                # body のみ対象。else 節は古い環境で実行されるので検査を続ける
                for stmt in node.body:
                    for child in ast.walk(stmt):
                        if hasattr(child, "lineno"):
                            guarded.add(child.lineno)
        return guarded

    def report(self, node, message):
        line = getattr(node, "lineno", 0)
        if line in self._guarded_lines:
            return
        snippet = self.lines[line - 1].strip() if 0 < line <= len(self.lines) else ""
        self.problems.append((line, message, snippet))

    # --- from __future__ / from dataclasses (3.7+) --------------------------
    def visit_ImportFrom(self, node):
        if node.module == "__future__":
            for alias in node.names:
                if alias.name == "annotations":
                    self.report(node, "from __future__ import annotations は Python 3.7+")
        elif node.module == "dataclasses":
            self.report(node, "dataclasses は Python 3.7+（標準ライブラリ）")
        self.generic_visit(node)

    # --- dataclasses (3.7+) -------------------------------------------------
    def visit_Import(self, node):
        for alias in node.names:
            if alias.name.split(".")[0] == "dataclasses":
                self.report(node, "dataclasses は Python 3.7+（標準ライブラリ）")
        self.generic_visit(node)

    # --- PEP 585 / PEP 604 の型注記 ----------------------------------------
    def visit_Subscript(self, node):
        value = node.value
        if isinstance(value, ast.Name) and value.id in BUILTIN_GENERICS:
            self.report(node, "組み込み型の添字 %s[...] は Python 3.9+（typing を使う）" % value.id)
        self.generic_visit(node)

    def visit_BinOp(self, node):
        # `X | None` 形式の型注記。数値の OR と区別するため、両辺が型らしいときだけ報告する
        if isinstance(node.op, ast.BitOr) and self._looks_like_type(node.left) and self._looks_like_type(node.right):
            self.report(node, "PEP 604 の `X | Y` 型注記は Python 3.10+（Optional/Union を使う）")
        self.generic_visit(node)

    def _looks_like_type(self, node):
        if isinstance(node, ast.Constant) and node.value is None:
            return True
        if isinstance(node, ast.Name) and (node.id[:1].isupper() or node.id in BUILTIN_GENERICS):
            return True
        return isinstance(node, ast.Subscript) or isinstance(node, ast.BinOp)

    # --- 3.8+ の構文 --------------------------------------------------------
    def visit_NamedExpr(self, node):
        self.report(node, "セイウチ演算子 := は Python 3.8+")
        self.generic_visit(node)

    def visit_JoinedStr(self, node):
        # f"{x=}" 形式（3.8+）。通常の f-string は 3.6 でも使える
        for value in node.values:
            if isinstance(value, ast.FormattedValue) and getattr(value, "conversion", -1) == 114:
                src = self.lines[node.lineno - 1] if node.lineno <= len(self.lines) else ""
                if "=}" in src:
                    self.report(node, "f-string の自己文書化 {x=} は Python 3.8+")
        self.generic_visit(node)

    # --- 新しい API ---------------------------------------------------------
    def visit_Call(self, node):
        for kw in node.keywords:
            if kw.arg in NEW_KWARGS:
                self.report(node, NEW_KWARGS[kw.arg])
        if isinstance(node.func, ast.Attribute) and node.func.attr in NEW_ATTRS:
            self.report(node, NEW_ATTRS[node.func.attr])
        self.generic_visit(node)


def check_file(path):
    with open(path, "r") as f:
        source = f.read()
    try:
        tree = ast.parse(source, filename=path)
    except SyntaxError as exc:
        return [(exc.lineno or 0, "SyntaxError: %s" % exc.msg, "")]
    checker = Py36Checker(path, source)
    checker.visit(tree)
    return sorted(checker.problems)


def iter_python_files(targets):
    for target in targets:
        if os.path.isfile(target):
            yield target
        else:
            for root, _dirs, files in os.walk(target):
                if "__pycache__" in root:
                    continue
                for name in sorted(files):
                    if name.endswith(".py"):
                        yield os.path.join(root, name)


def main():
    targets = sys.argv[1:] or ["src/autotest"]
    total = 0
    checked = 0
    for path in iter_python_files(targets):
        checked += 1
        problems = check_file(path)
        if problems:
            print("%s" % path)
            for line, message, snippet in problems:
                print("  %4d: %s" % (line, message))
                if snippet:
                    print("        | %s" % snippet[:100])
            total += len(problems)

    print("")
    print("=" * 70)
    if total:
        print(" Python 3.6 非互換: %d 件 / %d ファイル" % (total, checked))
    else:
        print(" Python 3.6 互換: 問題なし（%d ファイル検査）" % checked)
    print("=" * 70)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
