import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import main_makegraph_performance as mgp

#  [FILE]ボタン押下時に呼び出し。選択したファイルのパスをテキストボックスに設定する。
def open_file_command(edit_box, file_type_list):
    file_path = filedialog.askopenfilename(filetypes = file_type_list)
    edit_box.delete(0, tk.END)
    edit_box.insert(tk.END, file_path)

# ファイル設定エリアのフレームを作成して返却する関数
def set_file_frame(parent_frame, label_text, file_type_list):
    file_frame = ttk.Frame(parent_frame)
    tk.Label(file_frame, text = label_text).pack(side = tk.LEFT)
    # テキストボックスの作成と配置
    file_frame.edit_box = tk.Entry(file_frame, width = 50)
    file_frame.edit_box.pack(side = tk.LEFT)
    # ボタンの作成と配置
    file_button = tk.Button(file_frame, text = 'FILE', width = 5\
        , command = lambda:open_file_command(file_frame.edit_box, file_type_list))
    file_button.pack(side = tk.LEFT)
    file_frame.pack()
    return file_frame

def set_config_frame(parent_frame):
    file_frame = ttk.Frame(parent_frame)
    tk.Label(file_frame, text='試験コンフィグ（はｎ）')



def set_radiobutton(parent_frame):
    main_frame = ttk.Frame(parent_frame)
    var = tk.StringVar()
    rb1 = tk.Radiobutton(main_frame, value='Pa', variable=var, text='Pa')
    rb2 = tk.Radiobutton(main_frame, value='Torr', variable=var, text='Torr')
    rb1.pack()
    rb2.pack()
    main_frame.pack()

def set_do_button(parent_frame, edit_box_frame):
    main_frame = ttk.Frame(parent_frame)
    do_button = tk.Button(main_frame, text='DO', command=lambda:print(edit_box_frame.edit_box.get()))
    do_button.pack()
    main_frame.pack()

# フレームを作成する関数を呼び出して配置
def set_main_frame(root_frame):
    # ファイル選択エリア作成（ファイルの拡張子を指定）
    file_frame = set_file_frame(root_frame, "ファイル", [('excelブック', '*.xlsx'), ('excelマクロ', '*.xlsm')])
    set_radiobutton(root_frame)

    set_do_button(root_frame, file_frame)

# メイン関数
if __name__ == '__main__':
    root = tk.Tk()
    root.title('Tkinter training')
    root.geometry("500x300")
    set_main_frame(root)
    end_button = tk.Button(root, text = 'END', width = 10, command = lambda:root.quit())
    end_button.pack()
    root.mainloop()