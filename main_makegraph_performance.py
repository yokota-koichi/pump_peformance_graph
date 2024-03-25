"""背側測定シートのデータを整理する"""
import mod_makegraph_performance as mod
import tkinter

if __name__ == '__main__':


    # 今回の試験コンフィグを記入してください．(グラフの凡例名になります)
    test_config = 'ver2, 追加工ハウジング'

    # SRGの単位（測定シートのG列に入力される値）を'Torr'か'Pa'か選んでください
    dim_srg = 'Pa'

    # ファイルパスを記入．このpyファイルからの相対パス．もしくは絶対パス．
    file_name = r"C:\Users\shimadzu\OneDrive - SHIMADZU\0_ykt\01_project\GA87-2258_ステータ4段目生成物堆積対策水平展開\01_5305\GA87-2258-04_ステータ4段目生成物堆積対策水平展開_5305_排気性能測定\02_測定シート\02_温調なし\背圧調整排速測定シート_5305_N2_温調なし_生成物対策ｽﾃｰﾀ.xlsm"

    # 測定シートのシート名
    sheet_name = 'Sheet1'

    mod.data_process(file_name, sheet_name, test_config, dim_srg)


