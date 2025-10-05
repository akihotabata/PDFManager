# PDFManager

## 概要
PDFManager は、PDF の結合・分割・ページ編集をローカル環境で安全に行えるデスクトップアプリケーションです。  
Python と PySide6 を用いて構築され、直感的な UI で誰でも簡単に操作できます。

---

## 主な機能
- 複数 PDF ファイルの結合（並び替え・除外対応）  
- 指定ページごとの分割  
- ページ削除・回転・抽出などの軽編集  
- プレビュー表示（PyMuPDF 対応）  
- 結合時にファイル名をブックマークとして付与  
- サブフォルダを含めた PDF 一括読み込み  
- 自動ソート（自然順・日付順・サイズ順）  
- 高DPI対応・日本語パス対応  
- 完全ローカル動作（クラウド送信なし）  

---

## 動作環境
- 対応OS：Windows 10 / 11  
- Python：3.9 以上  
- 依存パッケージ：PySide6, pypdf, PyMuPDF  
- ライセンス：MIT License  

---

## ディレクトリ構成

```
PDFManager/
├─ src/
│  └─ pdf_merger_app.py          # メインアプリケーション
├─ tools/
│  ├─ run.bat                    # 自動仮想環境構築＋起動スクリプト
│  ├─ run_portable.bat           # （任意）ポータブルPython向け
│  ├─ build_exe.bat              # EXE化手順用（配布対象外）
│  └─ requirements.txt           # 必要パッケージ一覧
├─ docs/
│  ├─ README_ja.md               # 日本語詳細ドキュメント
│  ├─ README_en.md               # 英語簡易版
│  ├─ EXE化手順書_Windows.md    # EXE化手順
│  ├─ ARCHITECTURE.md            # 設計メモおよび技術構成
│  └─ screenshots/               # スクリーンショット（UI 3種）
├─ .gitignore
├─ LICENSE
└─ README.md                     # このファイル（メイン）
```

---

## セットアップ手順（Windows）

### 1. 仮想環境を作成
```
python -m venv venv
venv\Scripts\activate
```

### 2. 依存パッケージをインストール
```
pip install -r tools\requirements.txt
```

### 3. アプリを起動
```
python src\pdf_merger_app.py
```
または
```
tools\run.bat
```

---

## EXE化（任意・手順書参照）
```
pyinstaller --noconsole --onefile --name "PDF結合ツール" --icon=icon.ico src/pdf_merger_app.py
```

詳細手順は「docs/EXE化手順書_Windows.md」を参照。


---

## ライセンス
このプロジェクトは MIT ライセンスの下で公開されています。
