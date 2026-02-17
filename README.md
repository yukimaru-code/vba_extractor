# vbaEXTRACTOR

ExcelファイルからVBAコードを抽出するツール。

## Setup
```bash
pip install -r requirements.txt
```

## Usage
```bash
python vbaEXTRACTOR.py
```

- `tkinterdnd2` が利用可能な環境では、ウィンドウにExcelファイルをドラッグ&ドロップして抽出できます。
- `tkinterdnd2` が使えない場合は、従来のファイル選択ダイアログで動作します。
- 起動画面のチェックボックスをONにすると、抽出元Excelと同じフォルダに `<抽出元ファイル名>_report.json` を出力します。
- JSONレポートには `timestamp` / `target_file` / `status` / `extracted_count` / `extracted_files` / `output_dir` / `message` を記録します。
