DOM_GET
=======

Web情報を取得するマクロのソースコード生成

#URL 
1.URLを入力する。  
2.新規にIEが開き、マウスカーソルの下の部品に半透明な赤い枠が表示されるので操作したい部品に対してクリックする。  
3.中断したい場合はCtrlキーを押下し、続行するには継続をクリックする。  
4.記録を終了する場合はダブルクリックする。  
5.Web情報ダイアログでNodeオブジェクトを選択する。  
6.オブジェクト選択でNodeオブジェクトの種別を選択する。  
7.オペレーション選択でNodeオブジェクトへの操作を選択する。  
8.操作に設定するものや入力するものがある場合は入力値で設定する。  
9.5-8を繰り返し操作するものを全て設定したらソース生成ボタンをクリックする。  
10.ソーステキストエリアでSleepを追加したい場合はその場所にカーソルを合わせて、Sleepを追加をクリックする。  
11.ソースをブックに追加するにはブックに追加をクリックしてブック名、追加するソースのモジュール名を入力する。  
   (ブック名を入力しなければ、新規のブックを作成して挿入する)

#Title  
1.現在開いているIEを選択する画面が表示されるので操作したいIEを選択する。  
2.あとはURLの場合と同じように操作する。  

#MainForm  
最初にMainFormを表示する

###継続取得  
前回に引き続き取得する。（同じobjIEを使用する)

###Nodeオブジェクト  
一括削除するにはクリアボタンをクリックする。  
指定の1件を削除するには右クリックして削除を選択。  

###Attribute  
全て表示されていないものはダブルクリックする。

###URL  
各NodeオブジェクトのURLが表示される。  
URLを確認するにはURLの確認ボタンをクリックする。
