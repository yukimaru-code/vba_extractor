import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from oletools.olevba import VBA_Parser
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}
SUPPORTED_EXTENSIONS = (".xlsm", ".xlsb", ".xls")


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


def is_supported_excel_file(file_path):
    return os.path.isfile(file_path) and file_path.lower().endswith(SUPPORTED_EXTENSIONS)


def parse_dnd_file_paths(root, data):
    """
    ドロップデータ(Tcl list形式)をファイルパスの配列に変換する。
    スペースを含むパスや中括弧付きパスに対応するため splitlist を使う。
    """
    return [p for p in root.tk.splitlist(data) if p]


def normalize_vba_code(vba_code):
    if isinstance(vba_code, bytes):
        return vba_code.decode("utf-8", errors="replace")
    return vba_code if isinstance(vba_code, str) else str(vba_code)


def build_no_macro_message(file_path):
    return (
        "VBAマクロを抽出できませんでした。\n"
        "マクロが存在しないか、ファイルが暗号化・保護・破損している可能性があります。\n"
        f"対象: {file_path}"
    )


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

        # 保存先フォルダ（必要になったときに作成）
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        dir_path = os.path.dirname(file_path)
        output_dir = os.path.join(dir_path, f"{base_name}")

        print(f"解析中: {file_path}")
        
        # VBAの解析と抽出
        vbap = VBA_Parser(file_path)
        try:
            count = 0
            used_names = set()
            # extract_macrosは (filename, stream_path, vba_filename, vba_code) を返す
            if vbap.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vbap.extract_macros():
                    if count == 0 and not os.path.exists(output_dir):
                        os.makedirs(output_dir)

                    # ファイル名を安全化し、重複しない保存先を作る
                    save_path = build_unique_save_path(output_dir, vba_filename, used_names)
                    
                    # ファイル書き出し (エンコーディングはutf-8推奨)
                    with open(save_path, 'w', encoding='utf-8') as f:
                        f.write(normalize_vba_code(vba_code))
                    
                    print(f"保存: {save_path}")
                    count += 1

                return True, f"{count} 個のファイルを抽出しました。\n保存先: {output_dir}"

            return False, build_no_macro_message(file_path)
        finally:
            vbap.close()

    except Exception as e:
        return (
            False,
            "抽出処理でエラーが発生しました。\n"
            "ファイルが暗号化・保護・破損している可能性があります。\n"
            f"詳細: {str(e)}"
        )

def main():
    def run_extraction(file_path):
        success, message = extract_vba_from_excel(file_path)
        if success:
            messagebox.showinfo("完了", message)
        else:
            messagebox.showerror("結果", message)

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="VBAを抽出したいExcelファイルを選択してください",
            filetypes=[("Excel Macro Files", "*.xlsm *.xlsb *.xls"), ("All Files", "*.*")]
        )
        if file_path:
            run_extraction(file_path)
        else:
            print("キャンセルされました")

    # tkinterdnd2 が使える場合はD&D対応UIを表示
    if TkinterDnD is not None and DND_FILES is not None:
        root = TkinterDnD.Tk()
        root.title("VBA Extractor")
        root.geometry("520x220")
        root.resizable(False, False)

        title = tk.Label(root, text="Excelファイルをここにドラッグ&ドロップ")
        title.pack(pady=(16, 8))

        drop_area = tk.Label(
            root,
            text="Drop Here",
            relief="groove",
            bd=2,
            width=52,
            height=5
        )
        drop_area.pack(padx=16, pady=8, fill="x")

        hint = tk.Label(root, text="対応形式: .xlsm / .xlsb / .xls")
        hint.pack(pady=(4, 8))

        browse_btn = tk.Button(root, text="ファイル選択...", command=browse_file)
        browse_btn.pack(pady=(0, 12))

        def on_drop(event):
            paths = parse_dnd_file_paths(root, event.data)
            if not paths:
                messagebox.showerror("結果", "ドロップされたパスを取得できませんでした。")
                return

            target_paths = [p for p in paths if is_supported_excel_file(p)]
            if not target_paths:
                messagebox.showerror("結果", "対応していないファイル形式、またはファイルが存在しません。")
                return

            for file_path in target_paths:
                run_extraction(file_path)

        drop_area.drop_target_register(DND_FILES)
        drop_area.dnd_bind("<<Drop>>", on_drop)
        root.mainloop()
        return

    # フォールバック: 従来のファイル選択ダイアログ
    root = tk.Tk()
    root.withdraw()
    browse_file()

if __name__ == "__main__":
    main()
