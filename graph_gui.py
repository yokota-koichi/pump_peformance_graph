import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import mod_make_graph as mgp




#  [FILE]ボタン押下時に呼び出し。選択したファイルのパスをテキストボックスに設定する。
def open_file_command(edit_box, file_type_list):
    file_path = filedialog.askopenfilename(filetypes = file_type_list)
    edit_box.delete(0, tk.END)
    edit_box.insert(tk.END, file_path)

# ファイル設定エリアのフレームを作成して返却する関数
def set_file_frame(parent_frame, label_text, file_type_list):
    filepath_frame = ttk.Frame(parent_frame)
    tk.Label(filepath_frame, text = label_text).pack(side = tk.LEFT)
    # テキストボックスの作成と配置
    filepath_frame.edit_box = tk.Entry(filepath_frame, width = 50)
    filepath_frame.edit_box.pack(side = tk.LEFT)
    # ボタンの作成と配置
    file_button = tk.Button(filepath_frame, text = 'FILE', width = 5\
        , command = lambda:open_file_command(filepath_frame.edit_box, file_type_list))
    file_button.pack(side = tk.LEFT)
    filepath_frame.pack(pady=10)
    return filepath_frame

def set_config_frame(parent_frame):
    config_frame = ttk.Frame(parent_frame)
    tk.Label(config_frame, text='[任意] 試験コンフィグを入力（グラフの系列名になる. 例：新設計ハウジング）').pack()
    config_frame.edit_box = tk.Entry(config_frame, width=40)
    config_frame.edit_box.pack()
    config_frame.pack()
    return config_frame

def set_sheetname_frame(parent_frame):
    sheetname_frame = ttk.Frame(parent_frame)
    tk.Label(sheetname_frame, text='測定データのシート名を入力').pack()
    sheetname_frame.edit_box = tk.Entry(sheetname_frame, width=40)
    sheetname_frame.edit_box.insert(0,'Sheet1')
    sheetname_frame.edit_box.pack()
    sheetname_frame.pack()
    return sheetname_frame

def set_dim_radiobutton(parent_frame):
    dim_p_frame = ttk.Frame(parent_frame)
    dim_p_frame.rb = tk.StringVar(value='Pa')
    tk.Label(dim_p_frame,text='SRGの出力値の単位を選んでください．').pack()
    rb1 = tk.Radiobutton(dim_p_frame, value='Pa', variable=dim_p_frame.rb, text='Pa')
    rb2 = tk.Radiobutton(dim_p_frame, value='Torr', variable=dim_p_frame.rb, text='Torr')
    rb1.pack()
    rb2.pack()
    dim_p_frame.pack()
    return dim_p_frame

def set_cautionlabel_frame(parent_frame):
    cautionlabel_frame = ttk.Frame(parent_frame)
    tk.Label(cautionlabel_frame, text='注意：背圧値が空欄かどうかで排速測定か背圧特性を判断しています．\nそのため測定シートのF列は').pack()
    tk.Label(cautionlabel_frame,text='排速測定の時は”空欄”\n背圧特性の引き切りは”0”を記入してください', font=('',10,'bold')).pack(pady=5)
    tk.Label(cautionlabel_frame,text='を記入してください．').pack()
    cautionlabel_frame.pack()



def get_parameter(filepath_frame, dim_p_frame, config_frame, sheetname_frame):
    file_path = filepath_frame.edit_box.get()
    dim_pressure = dim_p_frame.rb.get()
    config = config_frame.edit_box.get()
    sheetname = sheetname_frame.edit_box.get()
    return [file_path, dim_pressure, config, sheetname]

def make_graph(filepath_frame, dim_p_frame, config_frame, sheetname_frame):

    parameter = get_parameter(filepath_frame, dim_p_frame, config_frame, sheetname_frame)

    # ファイルパスを記入．このpyファイルからの相対パス．もしくは絶対パス．
    file_name = parameter[0]

    if file_name == '':
        messagebox.showerror('エラー!','ファイルを選択してください')
    # 今回の試験コンフィグを記入してください．(グラフの凡例名になります)
    test_config = parameter[2]

    # SRGの単位（測定シートのG列に入力される値）を'Torr'か'Pa'か選んでください
    dim_srg = parameter[1]

    # 測定シートのシート名
    sheet_name = parameter[3]

    try:
        pspq, backpressure = mgp.data_process(file_name, sheet_name, test_config, dim_srg)

    except Exception as e:
        print(e.__class__.__name__)
        print(e.args)
        print(e)
        print(f"{e.__class__.__name__}: {e}")


    if (pspq == True) and ( backpressure == True):
        messagebox.showinfo('確認', 'PS，PQ曲線，背圧特性のグラフを作成しました．')
    elif pspq == True:
        messagebox.showinfo('確認', 'PS，PQ曲線を作成しました．')
    elif backpressure == True:
        messagebox.showinfo('確認', '背圧特性のグラフを作成しました．')



def set_do_button(parent_frame, filepath_frame, dim_p_frame, config_frame, sheetname_frame):
    main_frame = ttk.Frame(parent_frame)
    do_button = tk.Button(main_frame, text='DO', command=lambda : make_graph(filepath_frame, dim_p_frame, config_frame, sheetname_frame))
    do_button.pack(pady=15)
    main_frame.pack()

# フレームを作成する関数を呼び出して配置
def set_main_frame(root_frame):
    # ファイル選択エリア作成（ファイルの拡張子を指定）
    filepath_frame = set_file_frame(root_frame, "ファイル", [('excelファイル', '*.xlsx;*.xlsm;*.xls')])
    dim_p_frame = set_dim_radiobutton(root_frame)
    config_frame = set_config_frame(root_frame)
    sheetname_frame = set_sheetname_frame(root_frame)
    set_do_button(root_frame, filepath_frame, dim_p_frame, config_frame, sheetname_frame)

# メイン関数
if __name__ == '__main__':
    root = tk.Tk()
    root.title('PQ,PS曲線，背圧特性グラフ作成')
    root.geometry("600x500")
    set_main_frame(root)
    end_button = tk.Button(root, text = 'END', width = 10, command = lambda:root.quit())
    end_button.pack(pady=20)
    set_cautionlabel_frame(root)
    root.mainloop()