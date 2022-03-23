# Ruled-lines-on-Excel--spreadsheet-for-SAS

This macro accesses an Excel sheet from within SAS and draws a ruled line.
The macro draws a line over the record when the data changes with the variable set by var.
The number of referenced variables is fixed at two; if there is only one, set the same value to var2.
If there are three or more, move the macro twice for two plus one.

The "dde_cell_format" macro is referenced from the following site.  
https://thank-you-blog.org/sas/sasmacro/dde_cell_format/

How to use  
%set_go_border(inds=ALL_DATA_632, var1=subjid, var2=jiki_n,sheet=6.3.2,offset = 5,StCol1=1,EndCol1=25,StCol2=2,EndCol2=25);

inds: name of dataset to be read.  
var1: Name of the first variable to be referenced.  
var2: Second referenced variable name.  
Sheet: Name of sheet to be written.  
offset: If the line to be written out is not the first line, this value adjusts the line to be written out.  
Stcol1: The first column number to be written out in var1.  
Endcol1: First column number to be written in var1.  
Stcol2: First column number to be written in var2.  
Endcol2: First column number to be written out in var2.  
Translated with www.DeepL.com/Translator (free version)

--------------------------------------------------------
SAS上からExcelのシートにアクセスして罫線を引くマクロです。
varで設定した変数でデータが変わった時にそのレコードの上に罫線を引きます。
参照する変数は2つで固定です。1つの場合は同じ値をvar2に設定してください。
3つ以上の場合はマクロを2回動かして2つ＋1つの用にしてください。

マクロ「dde_cell_format」は下記サイトを参照させていただきました。   
https://thank-you-blog.org/sas/sasmacro/dde_cell_format/

使用方法

%set_go_border(inds=ALL_DATA_632, var1=subjid, var2=jiki_n,sheet=6.3.2,offset = 5,StCol1=1,EndCol1=25,StCol2=2,EndCol2=25);

inds：読み込むデータセット名  
var1：１つめの参照する変数名  
var2：２つめの参照する変数名  
Sheet：書き込むシート名  
offset：書き出しを行う行が１行目ではない場合、この値で書き出し行を調整します。  
Stcol1：var1で書き出す最初の列番号  
Endcol1：var1で書き出す最初の列番号  
Stcol2：var2で書き出す最初の列番号    
Endcol2：var2で書き出す最初の列番号


