# makegraph_performance.py

## Overview
　排気性能測定で使用する測定シートから，PS，PQ曲線，背圧特性のグラフを作成するプログラムです．
　グラフ自体はエクセルの方が都合が良さそうなので，matplotlibではなく，pythonのエクセルAPIを使用して，グラフを描画しています．
　matplotlibに比べて，使い勝手が悪いのでグラフサイズなどはプログラムで作った後に調整してください．

## Requirement
- python

## Usage
###main_makegraph_performance.py
このpyファイルのみ変更します．
基本的には以下の4つのみを変更することになります．
####input
- test_config:
    - グラフの凡例名になる．何の試験かわかる文言をいれる．
- dim_srg:
    - SRGによって，測定シートのG列に記入される圧力の単位が異なる．TorrかPaとする．
- file_name:
    - 測定シートのファイルパスを入力．このpyファイルからの相対パス，もしくは絶対パスでもよい．
- sheet_name:
    - 測定シートのデータが記入されているシート名を入力．

####output
- PSPQというシートとbackpressureというシートが追加され，そこにグラフが作成されます．

###mod_makegraph_performance.py
モジュールを集めたpyファイルです．
特に変更する必要はありません．

上記のパラメータを変更したら，**main_makegraph_performance.pyを実行**してください．



## Features
- 基本的には測定シートのエクセルブックに新たなシートを追加して，そこにグラフを作成するのですが，エクセルAPIの仕様上，ブック開いているとできません．ブックを開いたまま，プログラムを実行した場合は，'new_book.xlsx'というブックファイルを新規作成して，そこにグラフを作成しています．

- 引ききりか背圧特性かどうかは，測定シートの背圧のセルが空欄かどうかで区別しています．なので，引ききりの時は背圧値は0ではなく，空欄でお願いします．

- 背圧特性のデータはsccmごとに空白行は空ける必要は特にないです．（はずです．）

## Reference
None
## Author
yokota-koichi

## Licence
None

