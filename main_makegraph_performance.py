"""背側測定シートのデータを整理する"""
import mod_makegraph_performance as mod
import tkinter

if __name__ == '__main__':
    print('======program start======\n\n')

    # 今回の試験コンフィグを記入してください．(グラフの凡例名になります)
    test_config = 'ver2, 追加工ハウジング'

    # SRGの単位（測定シートのG列に入力される値）を'Torr'か'Pa'か選んでください
    dim_srg = 'Torr' # or 'Pa'

    # ファイルパスを記入．このpyファイルからの相対パス．もしくは絶対パス．
    file_name = '背圧調整排速測定シートN2_7305_75c_nrw5＋stHUP+burn_231027_v2209.xlsm'

    # 測定シートのシート名
    sheet_name = 'Sheet1'

    mod.data_process(file_name, sheet_name, test_config, dim_srg)

    print('\n\n======program finished======')
