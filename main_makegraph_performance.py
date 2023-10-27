"""背側測定シートのデータを整理する"""
import mod_makegraph_performance as mod

print('======program start======\n\n')

# 今回の試験コンフィグを記入してください．(グラフの凡例名になります)
test_config = 'ver2, 追加工ハウジング'

# SRGの単位（測定シートのG列に入力される値）を'Torr'か'Pa'か選んでください
dim_srg = 'Torr' # or 'Pa'

# ファイルパスを記入．このpyファイルからの相対パス．もしくは絶対パス．
file_name = 'test.xlsx'

# 測定シートのシート名
sheet_name = 'Sheet1'

mod.data_process(file_name, sheet_name, test_config, dim_srg)

print('\n\n======program finished======')
