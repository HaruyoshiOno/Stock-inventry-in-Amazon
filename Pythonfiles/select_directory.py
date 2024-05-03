import os
import tkinter as tk
from tkinter import filedialog
import check_thunder_install
import tkinter.messagebox as msgbox

def select_directory_and_save(initial_dir):
    if check_thunder_install.install_thunderbird:

        # すでに保存されたディレクトリのファイルがあるかどうかを確認
        file_name = "amalost_directory.txt"
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        file_path = os.path.join(documents_path, file_name)
        
        if os.path.exists(file_path):
            # 保存されたディレクトリのパスを読み込む
            with open(file_path, "r") as file:
                directory_path = file.read().strip()
                print("保存されたディレクトリが読み込まれました:", directory_path)
        else:
            root = tk.Tk()
            root.withdraw()  # メインウィンドウを表示しないようにする

            # ディレクトリを選択
            directory_path = filedialog.askopenfilename(initialdir=initial_dir, title="Select file", filetypes=[("All files", "*.*")])
            if directory_path:
                print("選択されたディレクトリ:", directory_path)
                
                # パスをテキストファイルに保存
                with open(file_path, "w") as file:
                    file.write(directory_path)
                    
                print("ディレクトリパスがファイルに保存されました:", file_path)
            else:
                print("ディレクトリが選択されませんでした。")
                msgbox.showwarning("警告", "ディレクトリが選択されませんでした。")

                # ウィンドウの破棄
            root.destroy()
        
        return directory_path