# EXE化手順書（Windows）

このドキュメントは、PDFManager を Windows 向け単一の実行ファイル（.exe）にパッケージングする手順です。

## 前提
- Windows 10/11
- Python 3.10 以上
- ビルド元ディレクトリ直下に本リポジトリが展開されていること
- 任意：アプリ用アイコン `docs\icon.ico` を用意（無ければアイコンなしでビルド可能）

## 手順（最短）
1. 依存のインストール
   ```bat
   python -m venv venv
   venv\Scripts\activate
   pip install -r tools\requirements.txt
   pip install pyinstaller
   ```

2. ビルド（アイコンありの例）
   ```bat
   pyinstaller --noconsole --onefile --name "PDFManager" --icon docs\icon.ico ^
     --hidden-import fitz --hidden-import pymupdf ^
     src\pdf_manager_app.py
   ```

   - 出力先：`dist\PDFManager.exe`
   - コンソール非表示：`--noconsole`
   - 1ファイル：`--onefile`（初回起動がやや遅い代わりに配布が簡単）

3. 実行テスト  
   `dist\PDFManager.exe` をダブルクリックして起動確認してください。

## バッチで自動ビルド
`tools\build_exe.bat` をダブルクリックすると、仮想環境～ビルドまで自動実行します。

## 署名（任意）
企業配布などで SmartScreen の警告を抑えたい場合は、コードサイニング証明書を用意し、以下で署名します。  
（証明書の購入と維持には費用がかかります）
```bat
signtool sign /tr http://timestamp.sectigo.com /td sha256 /fd sha256 /a dist\PDFManager.exe
```

## よくあるエラーと対処
- **SmartScreen 警告**：未署名 EXE のため。発行元が不明と出ても「詳細情報→実行」で進められます。  
- **DLL が見つからない**：`venv` を作り直し、依存を再インストール。`pip install --upgrade pip` も実施。  
- **サイズが大きい**：`--onedir` 方式に変更すると起動が速くなる代わりにフォルダ配布になります。  
- **日本語パスで失敗**：プロジェクトパスを短い英数字に移して再ビルド。

## 注意
- PyInstaller の成果物は OS 依存です（Windows でビルドした EXE は Windows でのみ動作）。
- セキュリティ製品が初回実行時に検査する場合があります。配布前に複数端末で動作確認してください。
