import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from oletools.olevba import VBA_Parser

def extract_vba_from_excel(file_path):
    """
    指定されたExcelファイルからVBAマクロを抽出し、
    ファイル名と同名のフォルダに保存する関数
    """
    try:
        # ファイルの存在確認
        if not os.path.exists(file_path):
            print(f"エラー: ファイルが見つかりません - {file_path}")
            return False

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
        # extract_macrosは (filename, stream_path, vba_filename, vba_code) を返す
        if vbap.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vbap.extract_macros():
                # ファイル名に使えない文字を置換
                safe_filename = vba_filename.replace('/', '_').replace('\\', '_')
                
                # 標準モジュールやクラスモジュールの判別は難しいため、一律 .txt または .vba で保存
                save_path = os.path.join(output_dir, f"{safe_filename}.txt")
                
                # ファイル書き出し (エンコーディングはutf-8推奨)
                with open(save_path, 'w', encoding='utf-8') as f:
                    f.write(vba_code)
                
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