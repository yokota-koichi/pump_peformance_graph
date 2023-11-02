import win32com.client

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = True
file_path = r"C:\Users\shimadzu\OneDrive - SHIMADZU\ykt\03_document\2_my program\排気性能_グラフ作成\背圧調整排速測定シートN2_7305_75c_nrw5＋stHUP+burn_231027_v2209.xlsm"
wb = xl.Workbooks.Open(file_path)

# ブックを閉じる
wb.Close()
