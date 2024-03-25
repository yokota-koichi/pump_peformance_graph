import win32com.client
from collections import Counter
import sys
import pythoncom
from tkinter import messagebox





def data_sort(ws, start_row):
    # 最終行を取得．この場合，数値が入力されている最終行ではなく，書式設定されているセルの最終行．
    lastrow = ws.Cells(ws.Rows.Count, 3).End(-4162).Row


    # 引き切りのデータに該当する行番号を格納するリスト．
    list_rownum_pspq = []
    # 背圧特性のデータに該当する行番号を格納するリスト．
    list_rownum_backpressure = []
    for i in range(start_row, lastrow + 1):
        # まず，SRGの列にデータがない行と流量の列が0，もしくは空白の行を除外．
        if (ws.Cells(i,7).Value != None) and (ws.Cells(i,3).Value != (None or 0)):
            # 排気口圧力の列が空白の行の行番号をリストに格納．
            if ws.Cells(i,6).Value == None:
                list_rownum_pspq.append(i)
            # 空白ではない行の行番号をリストに格納．
            else: list_rownum_backpressure.append(i)
    pspq = None
    backpressure = None

    if list_rownum_pspq != []:
        # 引き切りのデータの，[流量，SRG値，VAT値，流速]のリストを作成．
        list_pspq = []
        for i in list_rownum_pspq:
            list_pspq.append([ws.Cells(i,3).Value, ws.Cells(i,7).Value, ws.Cells(i,11).Value,ws.Cells(i,12).Value])
        # 流量を昇順に並び替え
        list_pspq.sort()
        pspq = True

    if list_rownum_backpressure != []:
        # ここから背圧特性のデータ整理
        # C列からsccmを取得
        list_sccm = []
        for i in list_rownum_backpressure:
            list_sccm.append(round(ws.Cells(i,3).Value))

        # list_sccmの中から種類を抜き出す．（辞書｛sccm：個数｝を作成し，key()で種類を抜き出す）
        list_sccm_var = Counter(list_sccm).keys()

        # 背圧特性評価に必要なデータを辞書形式で保存．
        # key = 流量, Value = [排気口圧力，吸気口圧力]
        dict_backpressure = {}
        for i in sorted(list_sccm_var):
            dict_backpressure[i] = []
            for j in list_rownum_backpressure:
                if round(ws.Cells(j,3).Value) == i:
                    dict_backpressure[i].append([ws.Cells(j,11).Value, ws.Cells(j,7).Value])
            dict_backpressure[i].sort()
        backpressure = True



    if pspq == None and backpressure == None:
        messagebox.showerror('エラー！','グラフを作成するためのデータがありません!.')
        sys.exit()

    elif pspq == True and backpressure == None:
        return pspq, backpressure, list_pspq, None
    elif pspq == None and backpressure == True:
        return pspq, backpressure, None, dict_backpressure
    else:
        return pspq, backpressure, list_pspq, dict_backpressure


def write_pspq_data(ws, test_config, list_pspq, dim_srg):
    list_word = [['','',test_config,'','','',''],['Gas throughput', '', 'Inlet pressure', '', 'Foreline pressure', '', 'Pumping speed'], ['sccm', '', 'Torr', 'Pa', 'Torr', 'Pa', 'L/s']]

    if dim_srg == 'Torr':
        coef_mat = [1,1/133.32]
    elif dim_srg == 'Pa':
        coef_mat = [133.32, 1]
    else:
        messagebox.showerror('SRGの単位が正しく記入されていません')

    for i in range(len(list_word)):
        for j in range(len(list_word[0])):
            ws.Cells(i+1,j+1).Value= list_word[i][j]

    for i in range(len(list_pspq)):
        ws.Cells(i+4,1).Value = list_pspq[i][0]
        ws.Cells(i+4,3).Value=list_pspq[i][1]*coef_mat[0]
        ws.Cells(i+4,4).Value=list_pspq[i][1]*coef_mat[1]
        ws.Cells(i+4,5).Value=list_pspq[i][2]*coef_mat[0]
        ws.Cells(i+4,6).Value=list_pspq[i][2]*coef_mat[1]
        ws.Cells(i+4,7).Value=list_pspq[i][3]

def make_pspq_curve(ws, test_config):
    lastrow = ws.Cells(ws.Rows.Count, 3).End(-4162).Row
    # PS曲線用の散布図を作成
    ps = ws.Shapes.AddChart2(240,74).Chart
    ws.ChartObjects(1).Left = 100
    ws.ChartObjects(1).Top = 400
    ws.ChartObjects(1).Width = 400
    ws.ChartObjects(1).Height = 250
    ps.FullSeriesCollection(1).XValues = "=%s!$C$4:$C$%d" %(ws.Name, lastrow)
    ps.FullSeriesCollection(1).Values = '=%s!$G$4:$G$%d' %(ws.Name, lastrow)
    ps.FullSeriesCollection(1).Name = test_config
    ps.Hastitle = True
    ps.ChartTitle.Text='PS curve'
    ps.Axes(1).HasTitle = True
    ps.Axes(1).AxisTitle.Text = 'Inlet Pressure [Torr]'
    ps.Axes(2).HasTitle = True
    ps.Axes(2).AxisTitle.Text = 'Pumping speed [L/s]'
    ps.Axes(1).ScaleType = -4133
    ps.Axes(2).TickLabelPosition = -4134
    ps.Axes(1).HasMinorGridlines = True
    ps.HasLegend = True
    ps.Legend.IncludeInLayout = True





    # PS曲線用の散布図を作成
    pq = ws.Shapes.AddChart2(240,74).Chart
    ws.ChartObjects(2).Left = 550
    ws.ChartObjects(2).Top = 400
    ws.ChartObjects(2).Width = 400
    ws.ChartObjects(2).Height = 250
    pq.FullSeriesCollection(1).XValues = "=%s!$C$4:$C$25" %ws.Name
    pq.FullSeriesCollection(1).Values = '=%s!$A$4:$A$25' %ws.Name
    pq.FullSeriesCollection(1).Name = test_config
    pq.Hastitle = True
    pq.ChartTitle.Text='PQ curve'
    pq.Axes(1).ScaleType = -4133
    pq.Axes(2).TickLabelPosition = -4134
    pq.Axes(1).HasMinorGridlines = True
    pq.Axes(1).HasTitle = True
    pq.Axes(1).AxisTitle.Text = 'Inlet pressure [Torr]'
    pq.Axes(2).HasTitle = True
    pq.Axes(2).AxisTitle.Text = 'Gas throughput [sccm]'
    pq.HasLegend = True







def write_backpressure_data(ws, test_config, dict_backpressure, dim_srg):
    if dim_srg == 'Torr':
        coef_mat = [1,1/133.32]
    elif dim_srg == 'Pa':
        coef_mat = [133.32, 1]
    else:
        messagebox('エラー！','SRGの単位が選択されていません')

    sccm_list = list(dict_backpressure.keys())
    ws.Cells(1,1).Value= test_config
    for i in range(len(sccm_list)):
        ws.Cells(2,2*i+1).Value=sccm_list[i]
        ws.Cells(3,2*i+1).Value='Outlet pressure\n[Torr]'
        ws.Cells(3,2*i+2).Value='Inlet pressure\n[Torr]'
        for j in range(len(dict_backpressure[sccm_list[i]])):
            ws.Cells(j+4,2*i+1).Value=dict_backpressure[sccm_list[i]][j][0]
            ws.Cells(j+4,2*i+2).Value=dict_backpressure[sccm_list[i]][j][1]


"""
背圧特性のグラフを作成するモジュール

第1引数：ワークシート
第2引数：試験コンフィグ
第3引数：背圧特性のデータを格納している辞書
"""
def make_backpressure_curve(ws, test_config, dict_backpressure):

    # sccmの種類のリストで取得
    sccm_list = list(dict_backpressure.keys())
    # sccmの種類の数を取得
    num_series = len(sccm_list)

    ws.Shapes.AddChart2(240,74).Chart
    chart = ws.ChartObjects(1).Chart
    ws.ChartObjects(1).Left = 100
    ws.ChartObjects(1).Top = 200
    ws.ChartObjects(1).Width = 500
    ws.ChartObjects(1).Height = 500
    chart.ChartArea.ClearContents()



    for i in range(num_series):
        chart.SeriesCollection().NewSeries()
        chart.FullSeriesCollection(i+1).XValues = ws.Range(ws.Cells(4,2*i+1),ws.Cells(3+len(dict_backpressure[sccm_list[i]]),2*i+1))
        chart.FullSeriesCollection(i+1).Values = ws.Range(ws.Cells(4,2*i+2),ws.Cells(3+len(dict_backpressure[sccm_list[i]]),2*i+2))
        chart.FullSeriesCollection(i+1).Name = test_config + '-%ssccm' %sccm_list[i]

    chart.Hastitle = True
    chart.ChartTitle.Text='背圧特性'
    chart.Axes(1).ScaleType = -4133
    chart.Axes(2).ScaleType = -4133
    chart.Axes(1).HasMinorGridlines = True
    chart.Axes(2).HasMinorGridlines = True
    chart.Axes(1).HasTitle = True
    chart.Axes(1).AxisTitle.Text = 'Outlet pressure [Pa]'
    chart.Axes(2).HasTitle = True
    chart.Axes(2).AxisTitle.Text = 'Inlet pressure [sccm]'
    chart.Axes(1).TickLabelPosition = -4134
    chart.Axes(2).TickLabelPosition = -4134
    chart.HasLegend = True
    chart.Legend.IncludeInLayout = True
    chart.Legend.Position = -4152






def data_process(file_name,sheet_name,test_config, dim_srg):
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Visible = False
    wb = xl.Workbooks.Open(file_name)
    ws = wb.Worksheets(sheet_name)

    pspq, backpressure, list_pspq, dict_backpressure = data_sort(ws, 18)

    #データ整理とグラフを作るモジュールを呼び出す．
    if pspq == True:
        # PQPSシートが既に存在していたら連番のシート名（Sheet5など）のシートに作成する
        ws_pqps = wb.Worksheets.Add(After=ws)
        try:
            ws_pqps.Name = 'PQPS'
        except pythoncom.com_error:
            None


        write_pspq_data(ws_pqps, test_config, list_pspq, dim_srg)
        make_pspq_curve(ws_pqps, test_config)


    if backpressure == True:
        ws_backpressure = wb.Worksheets.Add(After=ws)
        try:
            ws_backpressure.Name = 'backpressure'
        except pythoncom.com_error:
            None

        write_backpressure_data(ws_backpressure, test_config, dict_backpressure,dim_srg)
        make_backpressure_curve(ws_backpressure, test_config, dict_backpressure)

    xl.Visible = True

    return pspq, backpressure




if __name__ == '__main__':


    # 今回の試験コンフィグを記入してください．(グラフの凡例名になります)
    test_config = ''

    # SRGの単位（測定シートのG列に入力される値）を'Torr'か'Pa'か選んでください
    dim_srg = 'Torr' # or 'Pa'

    # ファイルパスを記入．このpyファイルからの相対パス．もしくは絶対パス．
    file_name = r"C:\Users\shimadzu\OneDrive - SHIMADZU\ykt\03_my program\python\pump_performance\pywin\背圧調整排速測定シート_1704LMF(T1)_Ar_20231226.xlsm"
    # 測定シートのシート名
    sheet_name = 'Sheet1'

    data_process(file_name, sheet_name, test_config, dim_srg)