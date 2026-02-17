import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from oletools.olevba import VBA_Parser

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


def sanitize_filename(name, default_name="module", max_length=120):
    """Windowsでも安全に保存できるファイル名に正規化する。"""
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name))
    safe = safe.strip().rstrip(".")

    if not safe:
        safe = default_name

    # 予約名は拡張子の有無に関わらず先頭名で判定
    if safe.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        safe = f"_{safe}"

    if len(safe) > max_length:
        safe = safe[:max_length].rstrip(" .")

    return safe or default_name


def build_unique_save_path(output_dir, raw_name, used_names):
    """重複を避けた保存先パスを返す。"""
    base_name = sanitize_filename(raw_name)
    candidate = base_name
    index = 1

    while candidate.lower() in used_names:
        suffix = f"_{index}"
        allowed = max(1, 120 - len(suffix))
        candidate = f"{base_name[:allowed]}{suffix}"
        candidate = candidate.rstrip(" .")
        index += 1

    used_names.add(candidate.lower())
    return os.path.join(output_dir, f"{candidate}.txt")


def extract_vba_from_excel(file_path):
    """
    指定されたExcelファイルからVBAマクロを抽出し、
    ファイル名と同名のフォルダに保存する関数
    """
    try:
        # ファイルの存在確認
        if not os.path.exists(file_path):
            print(f"エラー: ファイルが見つかりません - {file_path}")
            return False, f"ファイルが見つかりません: {file_path}"

        # 保存先フォルダの作成（Excelファイル名）
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        dir_path = os.path.dirname(file_path)
        output_dir = os.path.join(dir_path, f"{base_name}")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        print(f"解析中: {file_path}")
        
        # VBAの解析と抽出
        vbap = VBA_Parser(file_path)
        
        count = 0
        used_names = set()
        # extract_macrosは (filename, stream_path, vba_filename, vba_code) を返す
        if vbap.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vbap.extract_macros():
                # ファイル名を安全化し、重複しない保存先を作る
                save_path = build_unique_save_path(output_dir, vba_filename, used_names)
                
                # ファイル書き出し (エンコーディングはutf-8推奨)
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(vba_code if isinstance(vba_code, str) else str(vba_code))
                
                print(f"保存: {save_path}")
                count += 1
            
            vbap.close()
            return True, f"{count} 個のファイルを抽出しました。\n保存先: {output_dir}"
        else:
            vbap.close()
            return False, "VBAマクロが見つかりませんでした。"

    except Exception as e:
        return False, f"エラーが発生しました: {str(e)}"

def main():
    # GUIウィンドウを表示しない設定
    root = tk.Tk()
    root.withdraw()

    # ファイル選択ダイアログを開く
    file_path = filedialog.askopenfilename(
        title="VBAを抽出したいExcelファイルを選択してください",
        filetypes=[("Excel Macro Files", "*.xlsm *.xlsb *.xls"), ("All Files", "*.*")]
    )

    if file_path:
        success, message = extract_vba_from_excel(file_path)
        if success:
            messagebox.showinfo("完了", message)
        else:
            messagebox.showerror("結果", message)
    else:
        print("キャンセルされました")

if __name__ == "__main__":
    main()
