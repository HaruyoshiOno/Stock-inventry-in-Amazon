import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
import icon

from_date = None
to_date = None
statusbar = None
root = None
var = None
finished_progress = False

def view_gui():
    global root
    global from_cal, to_cal
    global statusbar
    global var
    global finished_progress



    if root is not None:
        root.geometry('400x90')
        return

    root = tk.Tk()
    photo = icon.get_photo_image4icon()
    root.iconphoto(False, photo)
    root.title('Amazon欠品対応ツール')
    root.geometry('400x0')
    var = tk.IntVar()

    def pickDate_from(event):
        global from_date
        from_date = event.widget.get_date()
        #print(event.widget.get_date())

    def pickDate_to(event):
        global to_date
        to_date = event.widget.get_date()
        #print(event.widget.get_date())

    #開始日付選択
    label1 = tk.Label(root, text='開始日:')
    label1.grid(row=1, column=0)
    from_cal_frame = tk.Frame(root)
    from_cal_frame.grid(row=1, column=1,columnspan=4 ,padx=0, pady=5, sticky='SW')
    from_cal = DateEntry(from_cal_frame, year=2024, locale='ja_JP')
    from_cal.grid()
    from_cal.bind("<<DateEntrySelected>>", pickDate_from)

    #終了日付選択
    label2 = tk.Label(root, text='終了日:')
    label2.grid(row=1, column=4)
    to_cal_frame = tk.Frame(root)
    to_cal_frame.grid(row=1, column=5, columnspan=3 ,padx=0, pady=5, sticky='SW')
    to_cal = DateEntry(to_cal_frame, year=2024, locale='ja_JP')
    to_cal.grid()
    to_cal.bind("<<DateEntrySelected>>", pickDate_to)

    #実行ボタン
    from main import main_func 
    count_btn=tk.Button(root,text="実行", width=5 ,command=main_func)
    count_btn.grid(row=2, column=7, padx=15, pady=5, sticky='N')

    #確定的Progressbar
    pb=ttk.Progressbar(root,maximum=1000,mode="determinate",variable=var, length=300)
    pb.grid(row=2, column=0, padx=15, columnspan=7, pady=5, sticky='N')

    #ステータスバー
    import check_thunder_install
    mail_flg = 'outlook'    #デフォルトのメーラー
    if check_thunder_install.install_thunderbird:
        mail_flg = 'Thunderbird'
    statusbar = tk.Label(root, text = f'{mail_flg}: 開始日⇒終了日の欠品メールを取得します', bd = 1, relief = tk.SUNKEN, anchor = tk.W)
    statusbar.grid(row=3, column=0, columnspan=8, sticky='EW')

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
    return

def countUp(count):
    if var.get()<1000:
        var.set(var.get()+count)
        #print(var.get()/10,'%')
        root.update()
        import main
        if var.get()==1000 and finished_progress == True:
            from main import update_label_text
            update_label_text('完了')
            messagebox.showinfo("Info","処理が完了しました！")
            root.quit()

def on_closing():
    if messagebox.askokcancel("確認", "本当に閉じていいですか？"):
        root.quit()