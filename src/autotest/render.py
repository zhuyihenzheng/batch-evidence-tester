"""証跡画像の生成（コードで描画する方式）。

対象は「フォルダ一覧」「ログ」「ファイル内容」の 3 種類。
DB データは画像化せず Excel のネイティブ表として出す（excel.py 側の責務）。

実画面キャプチャではなくコード描画にしている理由:
  - 対話ログオンセッション不要 = タスクスケジューラ / CI から無人実行できる
  - 解像度・ウィンドウ位置・IME 状態に左右されず、同じ入力なら同じ画像になる
  - 長いログでも自動でページ分割でき、切れた証跡にならない
見た目は Explorer / テキストエディタのウィンドウに寄せ、証跡としての可読性を確保する。
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont

from .fsops import FileEntry
from .logs import LineClassifier

# --- 配色（Windows 11 のライトテーマに寄せる）---------------------------------
C_WINDOW_BG = (255, 255, 255)
C_TITLEBAR = (243, 243, 243)
C_TITLE_TEXT = (32, 32, 32)
C_ADDRESS_BG = (251, 251, 251)
C_ADDRESS_BORDER = (219, 219, 219)
C_HEADER_BG = (247, 247, 247)
C_HEADER_TEXT = (90, 90, 90)
C_ROW_ALT = (250, 250, 250)
C_TEXT = (26, 26, 26)
C_MUTED = (120, 120, 120)
C_BORDER = (222, 222, 222)
C_GUTTER_BG = (245, 245, 245)
C_ACCENT = (0, 95, 184)
C_ERROR = (185, 28, 28)
C_WARN = (180, 110, 0)

_font_cache: Dict[Tuple[str, int], ImageFont.FreeTypeFont] = {}


def text_width(draw: "ImageDraw.ImageDraw", text: str, font) -> float:
    """文字列の描画幅を返す。

    Pillow のバージョン差を吸収する:
      - Pillow 8.0+ : textlength（推奨。Pillow 10 で textsize は削除された）
      - Pillow 5〜7 : textsize（Anaconda 5.2 同梱の 5.1 はこちら）
    """
    if hasattr(draw, "textlength"):
        return draw.textlength(text, font=font)
    return draw.textsize(text, font=font)[0]


def resolve_font(candidates: List[str], size: int) -> ImageFont.FreeTypeFont:
    """候補から実在するフォントを探す。日本語が出せないと証跡として使えないため必須。"""
    for path in candidates:
        key = (path, size)
        if key in _font_cache:
            return _font_cache[key]
        p = Path(path)
        if p.exists():
            try:
                font = ImageFont.truetype(str(p), size)
                _font_cache[key] = font
                return font
            except OSError:
                continue
    return ImageFont.load_default()


class Renderer:
    """1 テストケース分の画像を描くレンダラ。"""

    def __init__(self, evidence_cfg: dict, out_dir: Path):
        self.cfg = evidence_cfg
        self.out_dir = Path(out_dir)
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.width = int(evidence_cfg.get("image_width", 1180))
        size = int(evidence_cfg.get("font_size", 14))
        candidates = list(evidence_cfg.get("font_candidates", []))
        self.font = resolve_font(candidates, size)
        self.font_bold = resolve_font(candidates, size + 1)
        self.font_small = resolve_font(candidates, max(10, size - 2))
        self.line_h = size + 10
        self._seq = 0
        self._classifier = LineClassifier(evidence_cfg)

    def _next_path(self, stem: str) -> Path:
        self._seq += 1
        safe = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in stem)
        return self.out_dir / f"{self._seq:02d}_{safe}.png"

    # -- 共通パーツ ---------------------------------------------------------
    def _draw_window_chrome(self, draw: ImageDraw.ImageDraw, title: str, subtitle: str) -> int:
        """タイトルバー + アドレスバーを描き、本文開始 Y 座標を返す。"""
        w = self.width
        draw.rectangle([0, 0, w, 36], fill=C_TITLEBAR)
        draw.text((14, 10), title, font=self.font_bold, fill=C_TITLE_TEXT)
        # 右上のウィンドウボタン。フォント依存の記号を避けて図形で描く
        bx, by = w - 88, 17
        draw.line([bx, by, bx + 10, by], fill=C_MUTED)                              # 最小化
        draw.rectangle([bx + 28, by - 5, bx + 38, by + 5], outline=C_MUTED)          # 最大化
        draw.line([bx + 56, by - 5, bx + 66, by + 5], fill=C_MUTED)                  # 閉じる
        draw.line([bx + 66, by - 5, bx + 56, by + 5], fill=C_MUTED)
        draw.line([0, 36, w, 36], fill=C_BORDER)

        draw.rectangle([12, 46, w - 12, 74], fill=C_ADDRESS_BG, outline=C_ADDRESS_BORDER)
        draw.text((22, 53), subtitle, font=self.font, fill=C_TEXT)
        return 86

    def _footer(self, draw: ImageDraw.ImageDraw, y: int, text: str) -> int:
        draw.line([0, y, self.width, y], fill=C_BORDER)
        draw.text((14, y + 8), text, font=self.font_small, fill=C_MUTED)
        return y + 30

    def _draw_icon(self, draw: ImageDraw.ImageDraw, x: int, y: int, is_dir: bool) -> None:
        """フォルダ / ファイルのアイコンを図形で描く（絵文字はフォント依存で豆腐化するため）。"""
        if is_dir:
            draw.polygon([(x, y + 12), (x, y + 2), (x + 6, y + 2), (x + 8, y + 4), (x + 15, y + 4), (x + 15, y + 12)],
                         fill=(255, 205, 100), outline=(214, 163, 60))
        else:
            draw.polygon([(x + 1, y), (x + 9, y), (x + 13, y + 4), (x + 13, y + 13), (x + 1, y + 13)],
                         fill=(255, 255, 255), outline=(150, 150, 150))
            draw.polygon([(x + 9, y), (x + 13, y + 4), (x + 9, y + 4)], fill=(220, 220, 220), outline=(150, 150, 150))

    def _ellipsize(self, draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> str:
        if text_width(draw, text, font=font) <= max_width:
            return text
        while text and text_width(draw, text + "...", font=font) > max_width:
            text = text[:-1]
        return text + "..."

    # -- フォルダ一覧 -------------------------------------------------------
    def folder_listing(
        self,
        directory: Path,
        entries: List[FileEntry],
        label: str,
        phase: str,
        max_entries: int = 40,
    ) -> Path:
        """Explorer の詳細表示に相当する一覧画像を生成する。"""
        shown = entries[:max_entries]
        overflow = len(entries) - len(shown)

        header_h = 86
        col_head_h = 30
        rows_h = max(1, len(shown) + (1 if overflow > 0 else 0) + (1 if not entries else 0)) * self.line_h
        height = header_h + col_head_h + rows_h + 44
        img = Image.new("RGB", (self.width, height), C_WINDOW_BG)
        draw = ImageDraw.Draw(img)

        y = self._draw_window_chrome(draw, f"{label}（{phase}）", str(directory))

        # 列レイアウト: 名前 / 更新日時 / 種類 / サイズ
        x_name, x_mod, x_kind, x_size = 20, self.width - 470, self.width - 300, self.width - 130
        draw.rectangle([0, y, self.width, y + col_head_h], fill=C_HEADER_BG)
        for x, text in ((x_name, "名前"), (x_mod, "更新日時"), (x_kind, "種類"), (x_size, "サイズ")):
            draw.text((x, y + 7), text, font=self.font_small, fill=C_HEADER_TEXT)
        draw.line([0, y + col_head_h, self.width, y + col_head_h], fill=C_BORDER)
        y += col_head_h

        if not entries:
            draw.text((x_name, y + 6), "（このフォルダーは空です）", font=self.font, fill=C_MUTED)
            y += self.line_h
        else:
            for i, e in enumerate(shown):
                if i % 2 == 1:
                    draw.rectangle([0, y, self.width, y + self.line_h], fill=C_ROW_ALT)
                self._draw_icon(draw, x_name, y + 5, e.is_dir)
                name = self._ellipsize(draw, e.name, self.font, x_mod - x_name - 40)
                draw.text((x_name + 22, y + 4), name, font=self.font, fill=C_TEXT)
                draw.text((x_mod, y + 4), e.modified.strftime("%Y/%m/%d %H:%M"), font=self.font, fill=C_TEXT)
                draw.text((x_kind, y + 4), self._ellipsize(draw, e.kind, self.font, 160), font=self.font, fill=C_TEXT)
                size_text = e.size_text
                draw.text((x_size + (110 - text_width(draw, size_text, font=self.font)), y + 4), size_text, font=self.font, fill=C_TEXT)
                y += self.line_h
            if overflow > 0:
                draw.text((x_name, y + 4), f"... 他 {overflow} 件", font=self.font, fill=C_MUTED)
                y += self.line_h

        # ファイルとフォルダを分けて出す。サブフォルダ（Backup 等）が
        # 1 件として数えられ「ファイルが 1 件ある」と誤読されるのを防ぐ
        n_dir = sum(1 for e in entries if e.is_dir)
        n_file = len(entries) - n_dir
        summary = f"ファイル {n_file} 件"
        if n_dir:
            summary += f" / フォルダ {n_dir} 件"
        self._footer(draw, y, f"{summary}  |  取得日時: {datetime.now():%Y-%m-%d %H:%M:%S}")

        path = self._next_path(f"folder_{label}_{phase}")
        img.save(path)
        return path

    # -- テキスト（ログ / ファイル内容 / コンソール出力）---------------------
    def text_page(
        self,
        title: str,
        subtitle: str,
        lines: List[str],
        stem: str,
        max_lines: int = 45,
        max_cols: int = 160,
        show_line_numbers: bool = True,
        start_line_no: int = 1,
        highlight_keywords: Optional[List[str]] = None,
    ) -> List[Path]:
        """テキストを画像化する。max_lines を超える場合は複数ページに分割する。"""
        if not lines:
            lines = ["（内容なし）"]

        pages: List[Path] = []
        total_pages = (len(lines) + max_lines - 1) // max_lines
        for page in range(total_pages):
            chunk = lines[page * max_lines : (page + 1) * max_lines]
            path = self._render_text_page(
                title=title,
                subtitle=subtitle,
                lines=chunk,
                stem=f"{stem}_p{page + 1}" if total_pages > 1 else stem,
                page_label=f"{page + 1} / {total_pages}" if total_pages > 1 else "",
                max_cols=max_cols,
                show_line_numbers=show_line_numbers,
                start_line_no=start_line_no + page * max_lines,
                highlight_keywords=highlight_keywords or [],
            )
            pages.append(path)
        return pages

    def _render_text_page(
        self,
        title: str,
        subtitle: str,
        lines: List[str],
        stem: str,
        page_label: str,
        max_cols: int,
        show_line_numbers: bool,
        start_line_no: int,
        highlight_keywords: List[str],
    ) -> Path:
        height = 86 + len(lines) * self.line_h + 44
        img = Image.new("RGB", (self.width, height), C_WINDOW_BG)
        draw = ImageDraw.Draw(img)

        y = self._draw_window_chrome(draw, title, subtitle)

        gutter_w = 62 if show_line_numbers else 12
        if show_line_numbers:
            draw.rectangle([0, y, gutter_w, height - 44], fill=C_GUTTER_BG)
            draw.line([gutter_w, y, gutter_w, height - 44], fill=C_BORDER)

        for i, raw in enumerate(lines):
            line = raw.expandtabs(4)
            if len(line) > max_cols:
                line = line[:max_cols] + " ..."
            level = self._classifier.classify(line)
            if level == "error":
                color = C_ERROR
            elif level == "warn":
                color = C_WARN
            elif any(k and k in line for k in highlight_keywords):
                color = C_ACCENT
            else:
                color = C_TEXT

            if show_line_numbers:
                no = str(start_line_no + i)
                draw.text((gutter_w - 10 - text_width(draw, no, font=self.font_small), y + 5), no, font=self.font_small, fill=C_MUTED)
            draw.text((gutter_w + 12, y + 3), line, font=self.font, fill=color)
            y += self.line_h

        footer = f"{len(lines)} 行  |  取得日時: {datetime.now():%Y-%m-%d %H:%M:%S}"
        if page_label:
            footer += f"  |  ページ {page_label}"
        self._footer(draw, y, footer)

        path = self._next_path(stem)
        img.save(path)
        return path
