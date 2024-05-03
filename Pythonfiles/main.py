import tkinter.messagebox as msgbox
import tkinter as tk
import os
from datetime import datetime
from GUI import view_gui
import GUI
from select_directory import select_directory_and_save

import check_thunder_install
from check_thunder_install import check_thunderbird

def main_func():

    #必要なフォルダがなければ作成する
    create_folder_if_not_exists(folder_path_import)
    create_folder_if_not_exists(folder_path_upload)

    from_date = GUI.from_cal.get()
    to_date = GUI.to_cal.get()
    print("開始日:", from_date)
    print("終了日:", to_date)

    try:
        from_date_object = datetime.strptime(from_date, "%Y/%m/%d").date()
        to_date_object = datetime.strptime(to_date, "%Y/%m/%d").date()
    except ValueError as e:
        print("日付文字列の形式が正しくありません:", e) 
        msgbox.showwarning("警告", "日付形式が完全ではありません。")
        return

    if from_date_object == "" or to_date_object == "":
        msgbox.showwarning("警告", "日付が選択されていません。")
        return

    if from_date_object > to_date_object:
        msgbox.showwarning("警告", "開始日が終了日よりも後になっています。")
        return

    try:
        update_label_text('メールを取得中 step:1/2...')
        if check_thunder_install.install_thunderbird:
            from MailGet_thunderbird import MailGet_func_thunder
            MailGet_func_thunder()
        else:
            from MailGet import MailGet_func
            MailGet_func()
        update_label_text('メールの取得に成功しました')
        print('メールの取得に成功しました step:1/2')
    except ZeroDivisionError:
        update_label_text(f'メールの取得に失敗しました')
        print('指定日付内に受信した欠品メールがありません')
        msgbox.showwarning("警告", '指定日付内に受信した欠品メールがありません')
        GUI.root.quit()
        return
    except PermissionError:
        update_label_text(f'メール取得ファイルの書き出しに失敗しました')
        print('importファイルを閉じてください')
        msgbox.showwarning("警告", f'PermissionError: import_{today}.xlsxを閉じてください')
        GUI.root.quit()
        return
    except FileNotFoundError:
        update_label_text(f'メールの取得に失敗しました')
        print('メールフォルダのディレクトリが設定されていません')
        msgbox.showwarning("警告", f'FileNotFoundError: メールファイルのディレクトリが設定されていません。プログラムを再起動して設定を行ってください。')
        GUI.root.quit()
        return
    except Exception as e:
        update_label_text(f'メールの取得に失敗しました')
        msgbox.showwarning("警告", f'メールの取得に失敗しました:{e}')
        print(f'メールの取得に失敗しました:{e}')
        GUI.root.quit()
        return

    try:
        update_label_text('アップロードファイルを出力中 step:2/2...')
        from FileExport import FileExport_func
        FileExport_func()
        print('ファイルの出力に成功しました')
    except PermissionError:
        update_label_text(f'アップロード用ファイルの書き出しに失敗しました')
        print('uploadファイルを閉じてください')
        msgbox.showwarning("警告", f'PermissionError: upload_{today}.xlsxを閉じてください')
        GUI.root.quit()
        return
    except Exception as e:
        update_label_text(f'ファイルの出力に失敗しました')
        msgbox.showwarning("警告", f'ファイルの出力に失敗しました:{e}')
        GUI.root.quit()
        print(f'ファイルの出力に失敗しました:{e}')

def update_label_text(new_text):
    GUI.statusbar.config(text=new_text)

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"フォルダ '{folder_path}' を作成しました。")
    else:
        print(f"フォルダ '{folder_path}' は既に存在します。")

# 必要なフォルダ
folder_path_import = "mail_import"
folder_path_upload = "アップロード用ファイル"

if GUI.root is not None:

    check_thunderbird()
    if check_thunder_install.install_thunderbird:
        
        # 初期ディレクトリを設定
        initial_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Thunderbird", "Profiles")

        # ディレクトリを選択し、保存されたディレクトリのパスを取得
        selected_directory = select_directory_and_save(initial_dir)
        # 選択されたディレクトリを使用する
        print("ディレクトリのパス:", selected_directory)

# メインウィンドウの表示
view_gui()

today = format(datetime.today(), '%Y%m%d')