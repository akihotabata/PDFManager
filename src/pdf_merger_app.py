# pdf_manager_app.py
"""
PDF整理ツール

- ローカル専用（外部送信なし）
- 3タブ：[結合][分割][ページ編集]
- ページ編集：削除／抽出／回転／挿入／複製 + 右側プレビュー（ズーム・フィット）

依存:
  pip install PySide6 pypdf pymupdf
"""
from __future__ import annotations

import os
import sys
import re
import io
import traceback
from dataclasses import dataclass
from typing import List, Tuple, Optional

from PySide6 import QtCore, QtWidgets
from PySide6.QtCore import Qt, Signal, QThread
from PySide6.QtGui import QAction, QImage, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QStatusBar,
    QTabWidget,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)

from pypdf import PdfReader, PdfWriter

# PyMuPDF（任意）。未導入でも他機能は動作。
try:
    import fitz  # PyMuPDF
    _FitzOK = True
except Exception:
    fitz = None
    _FitzOK = False


APP_TITLE = "PDF整理ツール"
VERSION = "v2.1.0"  # + preview pane


# ---------- helpers ----------
def natural_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.findall(r"\d+|\D+", s)]


def parse_ranges(text: str, total_pages: int) -> List[Tuple[int, int]]:
    rngs: List[Tuple[int, int]] = []
    t = (text or "").strip()
    if not t:
        return rngs
    for part in t.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start = max(1, min(total_pages, int(a)))
            except ValueError:
                continue
            try:
                end = max(1, min(total_pages, int(b)))
            except ValueError:
                end = start
            if start > end:
                start, end = end, start
            rngs.append((start, end))
        else:
            try:
                p = max(1, min(total_pages, int(part)))
                rngs.append((p, p))
            except ValueError:
                continue
    return rngs


def rotate_page_inplace(page, deg: int):
    try:
        page.rotate(deg)
    except Exception:
        if deg % 360 == 90:
            page.rotate_clockwise(90)
        elif deg % 360 == 180:
            page.rotate_clockwise(180)
        elif deg % 360 == 270:
            page.rotate_counter_clockwise(90)


@dataclass
class PdfItem:
    path: str
    size: int

    @property
    def display(self) -> str:
        base = os.path.basename(self.path)
        kb = f"{self.size/1024:.1f} KB"
        return f"{base}  —  {kb}"


# ---------- workers ----------
class MergeWorker(QThread):
    progress = Signal(int)
    message = Signal(str)
    finished_ok = Signal(str)
    finished_error = Signal(str)

    def __init__(self, items: List[PdfItem], out_path: str, add_bookmarks: bool, parent=None):
        super().__init__(parent)
        self.items = items
        self.out_path = out_path
        self.add_bookmarks = add_bookmarks

    def run(self):
        try:
            if not self.items:
                self.finished_error.emit("結合対象のPDFがありません。")
                return

            writer = PdfWriter()
            total = len(self.items)
            done = 0

            for item in self.items:
                self.message.emit(f"読み込み中: {item.path}")
                try:
                    reader = PdfReader(item.path, strict=False)
                    if reader.is_encrypted:
                        try:
                            reader.decrypt("")
                        except Exception:
                            self.message.emit(f"スキップ(暗号化): {item.path}")
                            done += 1
                            self.progress.emit(int(done/total*100))
                            continue

                    start_page_index = len(writer.pages)
                    for p in reader.pages:
                        writer.add_page(p)

                    if self.add_bookmarks:
                        try:
                            writer.add_outline_item(os.path.basename(item.path), start_page_index)
                        except Exception:
                            self.message.emit(f"ブックマーク追加失敗: {os.path.basename(item.path)}")

                except Exception as e:
                    self.message.emit(f"スキップ(エラー): {item.path}  —  {e}")
                finally:
                    done += 1
                    self.progress.emit(int(done/total*100))

            os.makedirs(os.path.dirname(self.out_path) or ".", exist_ok=True)
            with open(self.out_path, "wb") as f:
                writer.write(f)

            self.finished_ok.emit(self.out_path)
        except Exception:
            self.finished_error.emit(traceback.format_exc())


class SplitWorker(QThread):
    progress = Signal(int)
    message = Signal(str)
    finished_ok = Signal(str)
    finished_error = Signal(str)

    MODE_EACH = "each"
    MODE_CHUNK = "chunk"
    MODE_RANGES = "ranges"

    def __init__(self, src_path: str, out_dir: str, filename_prefix: str,
                 mode: str, chunk_size: int = 1, ranges_text: str = "",
                 pad: int = 3, start_index: int = 1, parent=None):
        super().__init__(parent)
        self.src_path = src_path
        self.out_dir = out_dir
        self.filename_prefix = filename_prefix or "split"
        self.mode = mode
        self.chunk_size = max(1, int(chunk_size))
        self.ranges_text = ranges_text
        self.pad = max(1, int(pad))
        self.start_index = int(start_index)

    def run(self):
        try:
            if not os.path.exists(self.src_path):
                self.finished_error.emit("分割元PDFが存在しません。")
                return
            os.makedirs(self.out_dir or ".", exist_ok=True)

            reader = PdfReader(self.src_path, strict=False)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    self.finished_error.emit("暗号化PDFは分割できません。パスワード解除後に再実行してください。")
                    return

            total = len(reader.pages)
            idx = self.start_index

            def write_range(s1: int, e1: int):
                nonlocal idx
                w = PdfWriter()
                for p in range(s1 - 1, e1):
                    w.add_page(reader.pages[p])
                out_path = os.path.join(self.out_dir, f"{self.filename_prefix}_{str(idx).zfill(self.pad)}.pdf")
                with open(out_path, "wb") as f:
                    w.write(f)
                self.message.emit(f"出力: {out_path} (p{s1}-{e1})")
                idx += 1

            if self.mode == self.MODE_EACH:
                for p in range(1, total + 1):
                    write_range(p, p)
                    self.progress.emit(int(p / total * 100))
            elif self.mode == self.MODE_CHUNK:
                chunks = (total + self.chunk_size - 1) // self.chunk_size
                done = 0
                for c in range(chunks):
                    s = c * self.chunk_size + 1
                    e = min(total, (c + 1) * self.chunk_size)
                    write_range(s, e)
                    done = e
                    self.progress.emit(int(done / total * 100))
            else:
                rs = parse_ranges(self.ranges_text, total)
                if not rs:
                    self.finished_error.emit("カスタム範囲の指定が不正です。例: 1-3,5,7-10")
                    return
                for i, (s, e) in enumerate(rs, start=1):
                    write_range(s, e)
                    self.progress.emit(int(i / len(rs) * 100))

            self.finished_ok.emit(self.out_dir)
        except Exception:
            self.finished_error.emit(traceback.format_exc())


# ---------- main window ----------
class PdfManagerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} {VERSION}")
        self.resize(1180, 820)
        self.setMinimumSize(960, 620)

        self._build_ui()
        self._build_toolbar()
        self._build_menubar()

        self.status = QStatusBar()
        self.setStatusBar(self.status)

    # --- UI scaffold with scroll areas ---
    def _build_ui(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Merge tab
        self.merge_inner = QWidget()
        self._build_merge_tab(self.merge_inner)
        self.merge_scroll = QScrollArea()
        self.merge_scroll.setWidgetResizable(True)
        self.merge_scroll.setWidget(self.merge_inner)
        self.tabs.addTab(self.merge_scroll, "結合")

        # Split tab
        self.split_inner = QWidget()
        self._build_split_tab(self.split_inner)
        self.split_scroll = QScrollArea()
        self.split_scroll.setWidgetResizable(True)
        self.split_scroll.setWidget(self.split_inner)
        self.tabs.addTab(self.split_scroll, "分割")

        # Edit tab
        self.edit_inner = QWidget()
        self._build_edit_tab(self.edit_inner)
        self.edit_scroll = QScrollArea()
        self.edit_scroll.setWidgetResizable(True)
        self.edit_scroll.setWidget(self.edit_inner)
        self.tabs.addTab(self.edit_scroll, "ページ編集")

    # --- Merge Tab ---
    def _build_merge_tab(self, host: QWidget):
        self.items: List[PdfItem] = []

        main = QGridLayout(host)
        main.setContentsMargins(12, 12, 12, 12)
        main.setHorizontalSpacing(12)
        main.setVerticalSpacing(12)

        # left list
        left_box = QGroupBox("PDF一覧")
        left_layout = QVBoxLayout(left_box)
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        left_layout.addWidget(self.list_widget)

        btn_row = QHBoxLayout()
        self.btn_up = QPushButton("↑ 上へ")
        self.btn_down = QPushButton("↓ 下へ")
        self.btn_top = QPushButton("⤒ 先頭へ")
        self.btn_bottom = QPushButton("⤓ 末尾へ")
        self.btn_remove = QPushButton("✖ 選択削除")
        self.btn_clear = QPushButton("🗑 すべてクリア")
        for b in (self.btn_up, self.btn_down, self.btn_top, self.btn_bottom, self.btn_remove, self.btn_clear):
            b.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            btn_row.addWidget(b)
        left_layout.addLayout(btn_row)

        # right controls
        right_box = QGroupBox("操作")
        right = QGridLayout(right_box)
        right.setHorizontalSpacing(8)
        right.setVerticalSpacing(10)
        right.setColumnStretch(0, 0)
        right.setColumnStretch(1, 1)
        right.setColumnStretch(2, 0)

        self.txt_folder = QLineEdit(); self.txt_folder.setPlaceholderText("結合元フォルダのパス…")
        self.txt_folder.setMinimumWidth(320)
        self.btn_browse = QPushButton("フォルダ選択…")
        self.chk_recursive = QCheckBox("サブフォルダも含める")
        right.addWidget(QLabel("結合元フォルダ"), 0, 0)
        right.addWidget(self.txt_folder, 0, 1)
        right.addWidget(self.btn_browse, 0, 2)
        right.addWidget(self.chk_recursive, 1, 1)

        self.cmb_sort = QComboBox()
        self.cmb_sort.addItems([
            "ファイル名（自然順・昇順）",
            "ファイル名（自然順・降順）",
            "更新日時（新しい順）",
            "更新日時（古い順）",
            "サイズ（大きい順）",
            "サイズ（小さい順）",
        ])
        self.btn_scan = QPushButton("PDFを読み込む")
        right.addWidget(QLabel("並び順"), 2, 0)
        right.addWidget(self.cmb_sort, 2, 1)
        right.addWidget(self.btn_scan, 2, 2)

        self.txt_out = QLineEdit(); self.txt_out.setPlaceholderText("出力ファイル名 (例: merged.pdf)")
        self.txt_out.setMinimumWidth(320)
        self.btn_out = QPushButton("保存先…")
        self.chk_bookmark = QCheckBox("各ファイル名をブックマークにする")
        self.chk_open = QCheckBox("完了後にファイルを開く"); self.chk_open.setChecked(True)
        right.addWidget(QLabel("出力先"), 3, 0)
        right.addWidget(self.txt_out, 3, 1)
        right.addWidget(self.btn_out, 3, 2)
        right.addWidget(self.chk_bookmark, 4, 1)
        right.addWidget(self.chk_open, 4, 2)

        self.btn_merge = QPushButton("▶ 結合を実行")
        self.btn_merge.setMinimumHeight(44)
        right.addWidget(self.btn_merge, 5, 1, 1, 2)

        self.prog_merge = QProgressBar(); self.prog_merge.setRange(0, 100)
        self.log_merge = QTextEdit(); self.log_merge.setReadOnly(True); self.log_merge.setMinimumHeight(140)
        right.addWidget(QLabel("進捗"), 6, 0)
        right.addWidget(self.prog_merge, 6, 1, 1, 2)
        right.addWidget(QLabel("ログ"), 7, 0)
        right.addWidget(self.log_merge, 7, 1, 1, 2)

        main.addWidget(left_box, 0, 0, 2, 1)
        main.addWidget(right_box, 0, 1, 2, 1)
        main.setColumnStretch(0, 1)
        main.setColumnStretch(1, 2)
        main.setRowStretch(1, 1)

        self.list_widget.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.list_widget.model().rowsMoved.connect(self._sync_model_from_list)

        self.btn_browse.clicked.connect(self.on_browse)
        self.btn_scan.clicked.connect(self.on_scan)
        self.btn_out.clicked.connect(self.on_select_out)
        self.btn_merge.clicked.connect(self.on_merge)
        self.btn_up.clicked.connect(lambda: self._move_selected(-1))
        self.btn_down.clicked.connect(lambda: self._move_selected(+1))
        self.btn_top.clicked.connect(self._move_top)
        self.btn_bottom.clicked.connect(self._move_bottom)
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_clear.clicked.connect(self._clear_list)

    # --- Split Tab ---
    def _build_split_tab(self, host: QWidget):
        s = QGridLayout(host)
        s.setContentsMargins(12, 12, 12, 12)
        s.setHorizontalSpacing(12)
        s.setVerticalSpacing(12)

        src_box = QGroupBox("分割元PDF")
        src_l = QGridLayout(src_box)
        src_l.setColumnStretch(0, 0)
        src_l.setColumnStretch(1, 1)
        src_l.setColumnStretch(2, 0)
        self.txt_src = QLineEdit(); self.txt_src.setPlaceholderText("分割するPDFファイル…")
        self.txt_src.setMinimumWidth(360)
        self.btn_src = QPushButton("ファイル選択…")
        self.lbl_pages = QLabel("ページ数: -")
        src_l.addWidget(QLabel("ファイル"), 0, 0)
        src_l.addWidget(self.txt_src, 0, 1)
        src_l.addWidget(self.btn_src, 0, 2)
        src_l.addWidget(self.lbl_pages, 1, 1)

        mode_box = QGroupBox("分割モード")
        mode_l = QGridLayout(mode_box)
        mode_l.setColumnStretch(0, 0)
        mode_l.setColumnStretch(1, 1)
        self.rb_each = QRadioButton("1ページずつ"); self.rb_each.setChecked(True)
        self.rb_chunk = QRadioButton("Nページごと")
        self.rb_ranges = QRadioButton("カスタム範囲 (例: 1-3,5,7-10)")
        self.spin_chunk = QSpinBox(); self.spin_chunk.setRange(1, 9999); self.spin_chunk.setValue(10); self.spin_chunk.setMinimumWidth(100)
        self.txt_ranges = QLineEdit(); self.txt_ranges.setPlaceholderText("例: 1-3,5,7-10")
        mode_l.addWidget(self.rb_each, 0, 0, 1, 2)
        mode_l.addWidget(self.rb_chunk, 1, 0)
        mode_l.addWidget(self.spin_chunk, 1, 1)
        mode_l.addWidget(self.rb_ranges, 2, 0)
        mode_l.addWidget(self.txt_ranges, 2, 1)

        out_box = QGroupBox("出力")
        out_l = QGridLayout(out_box)
        out_l.setColumnStretch(0, 0)
        out_l.setColumnStretch(1, 1)
        out_l.setColumnStretch(2, 0)
        self.txt_outdir = QLineEdit(); self.txt_outdir.setPlaceholderText("出力フォルダ…"); self.txt_outdir.setMinimumWidth(360)
        self.btn_outdir = QPushButton("フォルダ選択…")
        self.txt_prefix = QLineEdit(); self.txt_prefix.setPlaceholderText("ファイル名プレフィックス (例: part)")
        self.spin_pad = QSpinBox(); self.spin_pad.setRange(1, 6); self.spin_pad.setValue(3); self.spin_pad.setMinimumWidth(90)
        self.spin_start = QSpinBox(); self.spin_start.setRange(0, 999999); self.spin_start.setValue(1); self.spin_start.setMinimumWidth(90)
        out_l.addWidget(QLabel("出力フォルダ"), 0, 0)
        out_l.addWidget(self.txt_outdir, 0, 1)
        out_l.addWidget(self.btn_outdir, 0, 2)
        out_l.addWidget(QLabel("プレフィックス"), 1, 0)
        out_l.addWidget(self.txt_prefix, 1, 1)
        out_l.addWidget(QLabel("連番の桁数"), 2, 0)
        out_l.addWidget(self.spin_pad, 2, 1)
        out_l.addWidget(QLabel("開始番号"), 3, 0)
        out_l.addWidget(self.spin_start, 3, 1)

        self.btn_split = QPushButton("▶ 分割を実行"); self.btn_split.setMinimumHeight(44)
        self.prog_split = QProgressBar(); self.prog_split.setRange(0, 100)
        self.log_split = QTextEdit(); self.log_split.setReadOnly(True); self.log_split.setMinimumHeight(140)

        s.addWidget(src_box, 0, 0)
        s.addWidget(mode_box, 1, 0)
        s.addWidget(out_box, 2, 0)
        s.addWidget(self.btn_split, 3, 0)
        s.addWidget(QLabel("進捗"), 4, 0)
        s.addWidget(self.prog_split, 5, 0)
        s.addWidget(QLabel("ログ"), 6, 0)
        s.addWidget(self.log_split, 7, 0)
        s.setRowStretch(7, 1)

        self.btn_src.clicked.connect(self.on_select_src)
        self.btn_outdir.clicked.connect(self.on_select_outdir)
        self.btn_split.clicked.connect(self.on_split)
        self.txt_src.textChanged.connect(self._update_page_count)

    # --- Edit Tab (with preview) ---
    def _build_edit_tab(self, host: QWidget):
        grid = QGridLayout(host)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)

        # Source
        src_box = QGroupBox("編集対象PDF")
        src = QGridLayout(src_box)
        src.setColumnStretch(0, 0); src.setColumnStretch(1, 1); src.setColumnStretch(2, 0)
        self.txt_edit_src = QLineEdit(); self.txt_edit_src.setPlaceholderText("編集するPDFファイル…")
        self.txt_edit_src.setMinimumWidth(360)
        self.btn_edit_src = QPushButton("ファイル選択…")
        self.lbl_edit_pages = QLabel("ページ数: -")
        src.addWidget(QLabel("ファイル"), 0, 0); src.addWidget(self.txt_edit_src, 0, 1); src.addWidget(self.btn_edit_src, 0, 2)
        src.addWidget(self.lbl_edit_pages, 1, 1)

        # Left: page list + actions
        left_box = QGroupBox("ページ一覧（選択して操作）")
        ll = QVBoxLayout(left_box)
        self.list_edit_pages = QListWidget()
        self.list_edit_pages.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.list_edit_pages.setAlternatingRowColors(True)
        self.list_edit_pages.setMinimumHeight(260)
        ll.addWidget(self.list_edit_pages)

        btns = QHBoxLayout()
        self.btn_page_delete = QPushButton("選択ページを削除")
        self.btn_page_extract = QPushButton("選択ページを抽出…")
        self.btn_page_rot_l = QPushButton("左回転90°")
        self.btn_page_rot_r = QPushButton("右回転90°")
        self.btn_page_dup = QPushButton("選択ページを複製（末尾）")
        self.btn_page_insert = QPushButton("挿入…")
        for b in (self.btn_page_delete, self.btn_page_extract, self.btn_page_rot_l, self.btn_page_rot_r, self.btn_page_dup, self.btn_page_insert):
            btns.addWidget(b)
        ll.addLayout(btns)

        # Right: preview panel
        preview_box = QGroupBox("プレビュー")
        pv = QGridLayout(preview_box)
        self.preview_info = QLabel("" if _FitzOK else "PyMuPDF（pymupdf）未インストールのためプレビュー無効\n\n有効化: pip install pymupdf")
        self.preview_info.setStyleSheet("color: gray;")
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidgetResizable(True)
        self.preview_container = QWidget()
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_label = QLabel("ページを選択してください")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setBackgroundRole(self.preview_label.backgroundRole())
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_layout.addWidget(self.preview_label)
        self.preview_scroll.setWidget(self.preview_container)

        ctrl = QHBoxLayout()
        self.btn_zoom_out = QPushButton("−"); self.btn_zoom_in = QPushButton("＋")
        self.btn_fit_w = QPushButton("幅に合わせる"); self.btn_fit_p = QPushButton("全体表示")
        ctrl.addWidget(self.btn_zoom_out); ctrl.addWidget(self.btn_zoom_in)
        ctrl.addStretch(1)
        ctrl.addWidget(self.btn_fit_w); ctrl.addWidget(self.btn_fit_p)

        pv.addLayout(ctrl, 0, 0)
        pv.addWidget(self.preview_scroll, 1, 0)
        pv.addWidget(self.preview_info, 2, 0)

        # Save area + log
        save_box = QGroupBox("保存")
        sv = QHBoxLayout(save_box)
        self.btn_save_over = QPushButton("上書き保存…")
        self.btn_save_as = QPushButton("別名で保存…")
        sv.addWidget(self.btn_save_over); sv.addWidget(self.btn_save_as)

        self.prog_edit = QProgressBar(); self.prog_edit.setRange(0, 100)
        self.log_edit = QTextEdit(); self.log_edit.setReadOnly(True); self.log_edit.setMinimumHeight(100)

        grid.addWidget(src_box, 0, 0, 1, 2)
        grid.addWidget(left_box, 1, 0)
        grid.addWidget(preview_box, 1, 1)
        grid.addWidget(save_box, 2, 0, 1, 2)
        grid.addWidget(QLabel("進捗"), 3, 0, 1, 2)
        grid.addWidget(self.prog_edit, 4, 0, 1, 2)
        grid.addWidget(QLabel("ログ"), 5, 0, 1, 2)
        grid.addWidget(self.log_edit, 6, 0, 1, 2)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 2)
        grid.setRowStretch(1, 1)
        grid.setRowStretch(6, 1)

        # Edit state
        self._edit_pages: List = []          # PageObject list (pypdf)
        self._edit_src_reader: Optional[PdfReader] = None
        self._edit_src_path: Optional[str] = None

        # Preview state
        self._preview_zoom_mode = "fitw"  # "fitw" / "fitp" / "free"
        self._preview_scale = 1.0         # used when free zoom
        self._current_preview_index = -1

        # Signals
        self.btn_edit_src.clicked.connect(self.on_edit_select_src)
        self.txt_edit_src.textChanged.connect(self._edit_load_from_path)

        self.list_edit_pages.currentRowChanged.connect(self._on_edit_selection_changed)

        self.btn_page_delete.clicked.connect(self.on_edit_delete_pages)
        self.btn_page_extract.clicked.connect(self.on_edit_extract_pages)
        self.btn_page_rot_l.clicked.connect(lambda: self.on_edit_rotate(-90))
        self.btn_page_rot_r.clicked.connect(lambda: self.on_edit_rotate(+90))
        self.btn_page_dup.clicked.connect(self.on_edit_duplicate_pages)
        self.btn_page_insert.clicked.connect(self.on_edit_insert_pages)

        self.btn_save_over.clicked.connect(lambda: self.on_edit_save(overwrite=True))
        self.btn_save_as.clicked.connect(lambda: self.on_edit_save(overwrite=False))

        self.btn_zoom_in.clicked.connect(lambda: self._set_zoom("free", 1.15))
        self.btn_zoom_out.clicked.connect(lambda: self._set_zoom("free", 1/1.15))
        self.btn_fit_w.clicked.connect(lambda: self._set_zoom("fitw"))
        self.btn_fit_p.clicked.connect(lambda: self._set_zoom("fitp"))
        self.preview_scroll.viewport().installEventFilter(self)

    # --- toolbar / menubar ---
    def _build_toolbar(self):
        tb = QToolBar("Main")
        tb.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, tb)

        act_browse = QAction("結合:フォルダを開く", self); act_browse.triggered.connect(self.on_browse)
        act_scan = QAction("結合:読み込み", self); act_scan.triggered.connect(self.on_scan)
        act_merge = QAction("結合開始", self); act_merge.triggered.connect(self.on_merge)

        act_split = QAction("分割開始", self); act_split.triggered.connect(self.on_split)

        act_edit_open = QAction("編集:PDFを開く", self); act_edit_open.triggered.connect(self.on_edit_select_src)

        tb.addAction(act_browse); tb.addAction(act_scan); tb.addSeparator()
        tb.addAction(act_merge); tb.addSeparator()
        tb.addAction(act_split); tb.addSeparator()
        tb.addAction(act_edit_open)

    def _build_menubar(self):
        m = self.menuBar()
        file_menu = m.addMenu("ファイル")
        act_open = QAction("結合:フォルダを開く…", self); act_open.triggered.connect(self.on_browse)
        act_out = QAction("結合:出力先を指定…", self); act_out.triggered.connect(self.on_select_out)
        act_src = QAction("分割:PDFを開く…", self); act_src.triggered.connect(self.on_select_src)
        act_outdir = QAction("分割:出力フォルダ…", self); act_outdir.triggered.connect(self.on_select_outdir)
        act_edit_open = QAction("編集:PDFを開く…", self); act_edit_open.triggered.connect(self.on_edit_select_src)
        act_exit = QAction("終了", self); act_exit.triggered.connect(self.close)

        for a in (act_open, act_out, act_src, act_outdir, act_edit_open):
            file_menu.addAction(a)
        file_menu.addSeparator()
        file_menu.addAction(act_exit)

        help_menu = m.addMenu("🛈 ヘルプ")
        act_about = QAction("このアプリについて", self); act_about.triggered.connect(self.on_about)
        act_how = QAction("使い方", self); act_how.triggered.connect(self.on_howto)
        help_menu.addAction(act_about); help_menu.addAction(act_how)

    # --- common logs ---
    def log_msg_merge(self, s: str):
        self.log_merge.append(s); self.status.showMessage(s, 5000)

    def log_msg_split(self, s: str):
        self.log_split.append(s); self.status.showMessage(s, 5000)

    def log_msg_edit(self, s: str):
        self.log_edit.append(s); self.status.showMessage(s, 5000)

    # ---------- Merge helpers/slots ----------
    def _collect_pdfs(self, folder: str, recursive: bool) -> List[PdfItem]:
        paths: List[Tuple[str, os.stat_result]] = []
        if recursive:
            for root, _, files in os.walk(folder):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        p = os.path.join(root, fn)
                        try: st = os.stat(p)
                        except Exception: continue
                        paths.append((p, st))
        else:
            try:
                for fn in os.listdir(folder):
                    if fn.lower().endswith(".pdf"):
                        p = os.path.join(folder, fn)
                        try: st = os.stat(p)
                        except Exception: continue
                        paths.append((p, st))
            except FileNotFoundError:
                pass

        mode = self.cmb_sort.currentText()
        if "昇順" in mode and "ファイル名" in mode:
            paths.sort(key=lambda t: natural_key(os.path.basename(t[0])))
        elif "降順" in mode and "ファイル名" in mode:
            paths.sort(key=lambda t: natural_key(os.path.basename(t[0])), reverse=True)
        elif "新しい順" in mode:
            paths.sort(key=lambda t: t[1].st_mtime, reverse=True)
        elif "古い順" in mode:
            paths.sort(key=lambda t: t[1].st_mtime)
        elif "大きい順" in mode:
            paths.sort(key=lambda t: t[1].st_size, reverse=True)
        elif "小さい順" in mode:
            paths.sort(key=lambda t: t[1].st_size)

        return [PdfItem(path=p, size=st.st_size) for p, st in paths]

    def _refresh_list_widget(self):
        self.list_widget.clear()
        for it in self.items:
            item = QListWidgetItem(it.display)
            item.setData(Qt.UserRole, it.path)
            self.list_widget.addItem(item)
        self.status.showMessage(f"{len(self.items)} 件のPDFを読み込みました。", 4000)

    def _sync_model_from_list(self):
        new_paths = [self.list_widget.item(i).data(Qt.UserRole) for i in range(self.list_widget.count())]
        path_to_item = {it.path: it for it in self.items}
        self.items = [path_to_item[p] for p in new_paths if p in path_to_item]

    def _move_selected(self, delta: int):
        rows = sorted({i.row() for i in self.list_widget.selectedIndexes()})
        if not rows:
            return
        for row in rows:
            new_row = row + delta
            if 0 <= new_row < self.list_widget.count():
                item = self.list_widget.takeItem(row)
                self.list_widget.insertItem(new_row, item)
                self.list_widget.setCurrentRow(new_row)
        self._sync_model_from_list()

    def _move_top(self):
        rows = sorted({i.row() for i in self.list_widget.selectedIndexes()})
        if not rows: return
        for idx, row in enumerate(rows):
            item = self.list_widget.takeItem(row - idx)
            self.list_widget.insertItem(idx, item)
        self._sync_model_from_list()

    def _move_bottom(self):
        rows = sorted({i.row() for i in self.list_widget.selectedIndexes()}, reverse=True)
        if not rows: return
        for row in rows:
            item = self.list_widget.takeItem(row)
            self.list_widget.addItem(item)
        self._sync_model_from_list()

    def _remove_selected(self):
        rows = sorted({i.row() for i in self.list_widget.selectedIndexes()}, reverse=True)
        for row in rows:
            self.list_widget.takeItem(row)
        self._sync_model_from_list()

    def _clear_list(self):
        self.items.clear()
        self.list_widget.clear()

    def on_browse(self):
        d = QFileDialog.getExistingDirectory(self, "フォルダを選択")
        if d: self.txt_folder.setText(d)

    def on_scan(self):
        folder = self.txt_folder.text().strip()
        if not folder:
            QMessageBox.warning(self, APP_TITLE, "結合元フォルダを指定してください。"); return
        recursive = self.chk_recursive.isChecked()
        self.items = self._collect_pdfs(folder, recursive)
        self._refresh_list_widget()
        if not self.items:
            QMessageBox.information(self, APP_TITLE, "PDFが見つかりませんでした。")

    def on_select_out(self):
        suggest = self.txt_out.text().strip() or os.path.join(os.path.expanduser("~"), "merged.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "出力ファイルを指定", suggest, "PDF (*.pdf)")
        if path:
            if not path.lower().endswith(".pdf"): path += ".pdf"
            self.txt_out.setText(path)

    def on_merge(self):
        if not self.items:
            QMessageBox.warning(self, APP_TITLE, "結合対象のPDFが一覧にありません。"); return
        out = self.txt_out.text().strip()
        if not out:
            QMessageBox.warning(self, APP_TITLE, "出力先のファイル名を指定してください。"); return

        self.log_merge.clear(); self.prog_merge.setValue(0); self.btn_merge.setEnabled(False)

        self.merge_worker = MergeWorker(self.items[:], out, self.chk_bookmark.isChecked())
        self.merge_worker.progress.connect(self.prog_merge.setValue)
        self.merge_worker.message.connect(self.log_msg_merge)
        self.merge_worker.finished_ok.connect(self._on_merge_finished_ok)
        self.merge_worker.finished_error.connect(self._on_merge_finished_error)
        self.merge_worker.finished.connect(lambda: self.btn_merge.setEnabled(True))
        self.merge_worker.start()

    def _on_merge_finished_ok(self, out_path: str):
        self.log_msg_merge(f"完了: {out_path}")
        QMessageBox.information(self, APP_TITLE, f"PDFの結合が完了しました。\n\n出力: {out_path}")
        if hasattr(self, "chk_open") and self.chk_open.isChecked():
            try:
                if sys.platform.startswith("win"): os.startfile(out_path)  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    import subprocess; subprocess.Popen(["open", out_path])
                else:
                    import subprocess; subprocess.Popen(["xdg-open", out_path])
            except Exception: pass

    def _on_merge_finished_error(self, detail: str):
        self.log_msg_merge("エラーが発生しました。ログを確認してください。")
        QMessageBox.critical(self, APP_TITLE, "PDFの結合に失敗しました。\n\n詳細:\n" + detail)

    # ---------- Split slots ----------
    def on_select_src(self):
        path, _ = QFileDialog.getOpenFileName(self, "分割するPDFを選択", "", "PDF (*.pdf)")
        if path: self.txt_src.setText(path)

    def on_select_outdir(self):
        d = QFileDialog.getExistingDirectory(self, "出力フォルダを選択")
        if d: self.txt_outdir.setText(d)

    def _update_page_count(self):
        p = self.txt_src.text().strip()
        if not p or not os.path.exists(p):
            self.lbl_pages.setText("ページ数: -"); return
        try:
            r = PdfReader(p, strict=False)
            if r.is_encrypted:
                try: r.decrypt("")
                except Exception:
                    self.lbl_pages.setText("ページ数: 暗号化"); return
            self.lbl_pages.setText(f"ページ数: {len(r.pages)}")
        except Exception:
            self.lbl_pages.setText("ページ数: 取得失敗")

    def on_split(self):
        src = self.txt_src.text().strip()
        if not src:
            QMessageBox.warning(self, APP_TITLE, "分割するPDFファイルを指定してください。"); return
        outdir = self.txt_outdir.text().strip() or os.path.join(os.path.dirname(src), "split")
        prefix = self.txt_prefix.text().strip() or "part"

        if self.rb_each.isChecked():
            mode = SplitWorker.MODE_EACH; chunk = 1; ranges = ""
        elif self.rb_chunk.isChecked():
            mode = SplitWorker.MODE_CHUNK; chunk = self.spin_chunk.value(); ranges = ""
        else:
            mode = SplitWorker.MODE_RANGES; chunk = 1; ranges = self.txt_ranges.text().strip()

        pad = self.spin_pad.value()
        start = self.spin_start.value()

        self.log_split.clear(); self.prog_split.setValue(0); self.btn_split.setEnabled(False)

        self.split_worker = SplitWorker(src, outdir, prefix, mode, chunk, ranges, pad, start)
        self.split_worker.progress.connect(self.prog_split.setValue)
        self.split_worker.message.connect(self.log_msg_split)
        self.split_worker.finished_ok.connect(self._on_split_finished_ok)
        self.split_worker.finished_error.connect(self._on_split_finished_error)
        self.split_worker.finished.connect(lambda: self.btn_split.setEnabled(True))
        self.split_worker.start()

    def _on_split_finished_ok(self, out_dir: str):
        self.log_msg_split(f"分割完了: {out_dir}")
        QMessageBox.information(self, APP_TITLE, f"PDFの分割が完了しました。\n\n出力フォルダ: {out_dir}")
        try:
            if sys.platform.startswith("win"): os.startfile(out_dir)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess; subprocess.Popen(["open", out_dir])
            else:
                import subprocess; subprocess.Popen(["xdg-open", out_dir])
        except Exception: pass

    def _on_split_finished_error(self, detail: str):
        self.log_msg_split("エラーが発生しました。ログを確認してください。")
        QMessageBox.critical(self, APP_TITLE, "PDFの分割に失敗しました。\n\n詳細:\n" + detail)

    # ---------- Edit helpers ----------
    def _refresh_edit_list(self):
        self.list_edit_pages.clear()
        for i, _ in enumerate(self._edit_pages, start=1):
            self.list_edit_pages.addItem(f"ページ {i}")
        self.lbl_edit_pages.setText(f"ページ数: {len(self._edit_pages)}")

    def _edit_load_from_path(self):
        p = self.txt_edit_src.text().strip()
        if not p or not os.path.exists(p):
            self._edit_pages = []; self._edit_src_reader = None; self._edit_src_path = None
            self._refresh_edit_list()
            self._render_preview(-1)
            return
        try:
            r = PdfReader(p, strict=False)
            if r.is_encrypted:
                try: r.decrypt("")
                except Exception:
                    QMessageBox.warning(self, APP_TITLE, "暗号化PDFは読み込めません。"); return
            self._edit_src_reader = r
            self._edit_src_path = p
            self._edit_pages = [r.pages[i] for i in range(len(r.pages))]
            self._refresh_edit_list()
            self.log_msg_edit(f"読み込み: {p}")
            self.list_edit_pages.setCurrentRow(0)
        except Exception as e:
            self._edit_pages = []; self._edit_src_reader = None; self._edit_src_path = None
            self._refresh_edit_list()
            self._render_preview(-1)
            QMessageBox.critical(self, APP_TITLE, f"読み込みに失敗しました。\n\n{e}")

    def on_edit_select_src(self):
        path, _ = QFileDialog.getOpenFileName(self, "編集するPDFを選択", "", "PDF (*.pdf)")
        if path:
            self.txt_edit_src.setText(path)  # triggers load

    def _on_edit_selection_changed(self, row: int):
        self._current_preview_index = row
        self._render_preview(row)

    def on_edit_delete_pages(self):
        if not self._edit_pages:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        rows = sorted({i.row() for i in self.list_edit_pages.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, APP_TITLE, "削除するページを選択してください。"); return
        for r in rows:
            self._edit_pages.pop(r)
        self._refresh_edit_list()
        next_row = min(rows[-1], len(self._edit_pages)-1) if self._edit_pages else -1
        self.list_edit_pages.setCurrentRow(next_row)
        self.log_msg_edit(f"削除: {len(rows)}ページ")

    def on_edit_extract_pages(self):
        if not self._edit_pages:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        rows = sorted({i.row() for i in self.list_edit_pages.selectedIndexes()})
        if not rows:
            QMessageBox.information(self, APP_TITLE, "抽出するページを選択してください。"); return
        path, _ = QFileDialog.getSaveFileName(self, "抽出先ファイル名", os.path.join(os.path.expanduser("~"), "extract.pdf"), "PDF (*.pdf)")
        if not path: return
        if not path.lower().endswith(".pdf"): path += ".pdf"
        w = PdfWriter()
        for r in rows:
            w.add_page(self._edit_pages[r])
        with open(path, "wb") as f:
            w.write(f)
        self.log_msg_edit(f"抽出: {len(rows)}ページ -> {path}")
        QMessageBox.information(self, APP_TITLE, f"抽出しました。\n\n出力: {path}")

    def on_edit_rotate(self, deg: int):
        if not self._edit_pages:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        rows = sorted({i.row() for i in self.list_edit_pages.selectedIndexes()})
        if not rows:
            QMessageBox.information(self, APP_TITLE, "回転するページを選択してください。"); return
        for r in rows:
            rotate_page_inplace(self._edit_pages[r], deg)
        self._refresh_edit_list()
        # 再描画
        self._render_preview(self._current_preview_index)
        self.log_msg_edit(f"回転: {len(rows)}ページ ({'左' if deg<0 else '右'}90°)")

    def on_edit_duplicate_pages(self):
        if not self._edit_pages:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        rows = sorted({i.row() for i in self.list_edit_pages.selectedIndexes()})
        if not rows:
            QMessageBox.information(self, APP_TITLE, "複製するページを選択してください。"); return
        copies = [self._edit_pages[r] for r in rows]
        self._edit_pages.extend(copies)
        self._refresh_edit_list()
        self.log_msg_edit(f"複製: {len(rows)}ページ（末尾に追加）")

    def on_edit_insert_pages(self):
        if self._edit_pages is None:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        pos = 0
        if self.list_edit_pages.currentRow() >= 0:
            pos = self.list_edit_pages.currentRow()
        src_path, _ = QFileDialog.getOpenFileName(self, "挿入元PDFを選択", "", "PDF (*.pdf)")
        if not src_path: return
        try:
            r = PdfReader(src_path, strict=False)
            if r.is_encrypted:
                try: r.decrypt("")
                except Exception:
                    QMessageBox.warning(self, APP_TITLE, "暗号化PDFは挿入できません。"); return
            total = len(r.pages)
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"挿入元の読み込みに失敗しました。\n\n{e}")
            return
        ranges, ok = QtWidgets.QInputDialog.getText(self, "挿入範囲", f"挿入するページ範囲を指定（1-{total}）：例 1-3,5")
        if not ok or not ranges.strip(): return
        rs = parse_ranges(ranges, total)
        if not rs:
            QMessageBox.information(self, APP_TITLE, "範囲の指定が不正です。"); return
        pages_to_insert = []
        for s, e in rs:
            for i in range(s-1, e):
                pages_to_insert.append(r.pages[i])
        for offset, pg in enumerate(pages_to_insert):
            self._edit_pages.insert(pos + offset, pg)
        self._refresh_edit_list()
        self.list_edit_pages.setCurrentRow(pos)
        self.log_msg_edit(f"挿入: {len(pages_to_insert)}ページ（位置: {pos+1} の前）")

    def on_edit_save(self, overwrite: bool):
        if not self._edit_pages:
            QMessageBox.warning(self, APP_TITLE, "PDFを読み込んでください。"); return
        dest = None
        if overwrite and self._edit_src_path:
            suggest = self._edit_src_path
            dest, _ = QFileDialog.getSaveFileName(self, "上書き保存（別名推奨）", suggest, "PDF (*.pdf)")
        else:
            suggest = os.path.join(os.path.dirname(self._edit_src_path or os.path.expanduser("~")), "edited.pdf")
            dest, _ = QFileDialog.getSaveFileName(self, "別名で保存", suggest, "PDF (*.pdf)")
        if not dest: return
        if not dest.lower().endswith(".pdf"): dest += ".pdf"

        try:
            w = PdfWriter()
            total = len(self._edit_pages)
            for i, p in enumerate(self._edit_pages, start=1):
                w.add_page(p)
                if total:
                    self.prog_edit.setValue(int(i/total*100))
            with open(dest, "wb") as f:
                w.write(f)
            self.log_msg_edit(f"保存: {dest}")
            QMessageBox.information(self, APP_TITLE, f"保存しました。\n\n出力: {dest}")
        except Exception as e:
            QMessageBox.critical(self, APP_TITLE, f"保存に失敗しました。\n\n{e}")

    # ---------- Preview rendering ----------
    def _set_zoom(self, mode: str, mul: Optional[float] = None):
        if mode in ("fitw", "fitp"):
            self._preview_zoom_mode = mode
        else:
            self._preview_zoom_mode = "free"
            self._preview_scale *= (mul or 1.0)
            self._preview_scale = max(0.1, min(8.0, self._preview_scale))
        self._render_preview(self._current_preview_index)

    def eventFilter(self, obj, event):
        # レイアウトサイズ変更時にフィット系を再計算
        if obj is self.preview_scroll.viewport() and event.type() == QtCore.QEvent.Resize:
            if self._preview_zoom_mode in ("fitw", "fitp"):
                self._render_preview(self._current_preview_index)
        return super().eventFilter(obj, event)

    def _render_preview(self, row: int):
        if not _FitzOK:
            self.preview_label.setText("プレビューを有効にするには:\n\npip install pymupdf")
            return
        if row is None or row < 0 or row >= len(self._edit_pages):
            self.preview_label.setText("ページを選択してください")
            return

        # 現在のPageObjectだけで1ページPDFを作り、メモリ上でレンダリング
        try:
            buf = io.BytesIO()
            w = PdfWriter()
            w.add_page(self._edit_pages[row])
            w.write(buf); data = buf.getvalue(); buf.close()

            doc = fitz.open(stream=data, filetype="pdf")
            page = doc[0]

            # 目標スケール計算
            viewport = self.preview_scroll.viewport().size()
            vw, vh = max(1, viewport.width()-6), max(1, viewport.height()-6)
            pw, ph = page.rect.width, page.rect.height

            if self._preview_zoom_mode == "fitw":
                scale = vw / pw
            elif self._preview_zoom_mode == "fitp":
                scale = min(vw / pw, vh / ph)
            else:
                scale = self._preview_scale

            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat, alpha=True)
            fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
            self.preview_label.setPixmap(QPixmap.fromImage(img))
            self.preview_info.setText(f"ページ {row+1} / {len(self._edit_pages)}  |  {pix.width}×{pix.height}px  |  zoom={scale:.2f}x")

            doc.close()
        except Exception as e:
            self.preview_label.setText(f"プレビューに失敗しました。\n{e}")

    # --- about / howto ---
    def on_about(self):
        QMessageBox.information(
            self, "このアプリについて",
            f"{APP_TITLE} {VERSION}\n\n"
            "ローカル専用のPDF整理ツールです。\n"
            "・結合：フォルダ読み込み／並べ替え／ブックマーク付与\n"
            "・分割：1ページずつ／Nページごと／範囲指定\n"
            "・ページ編集：削除／抽出／回転／挿入／複製\n"
            "・プレビュー：選択ページを右側に表示、ズーム／フィット対応\n"
        )

    def on_howto(self):
        QMessageBox.information(
            self, "使い方",
            "【ページ編集のプレビュー】\n"
            "・左リストでページを選択すると右側にプレビュー\n"
            "・＋/− で拡大縮小、[幅に合わせる]/[全体表示]で自動調整\n\n"
            "【結合/分割】は従来どおりです。"
        )


def main():
    # High DPI
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("LocalTools")

    w = PdfManagerWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
