適当な場所に保存してFTE-EngineCEA.exeを起動してもらえれば動きます。
最初から入っているものは消さないようにお願いします。多分ソフトが動かなくなります。
上に書いてある固定：・・・はnasa-ceaで入力する時に固定で設定するやつです（googleドライブの参考資料参照）
このソフトはあくまで入力を手助けするもので、実際に計算するものではありません。計算はnasa-ceaが行います（FECA2m.exe）
計算したいO/Fと燃焼圧の数が多すぎるとnasa-ceaで全ての計算結果を出せません。
このソフトは動いたのに計算結果がおかしいって場合はどこかしら値がおかしいです。
数値は全て半角でお願いします。
おかしい点や要望などあれば、近藤恵休（LINE名えきゅー）まで連絡をお願いします。いっぱい働きます。

○使い方
・基本的には、値入力→inp出力→cea起動→Excel出力の流れ

・値入力
現時点ではO/F、燃焼圧共に範囲にしか対応していません。
計算したい値の開始から終了と、その間隔を指定したらそれに合った値で計算します。少数可（例　1 to 3 間隔 0.5だと 1, 1.5, 2, 2.5, 3を計算します）
反応物質の指定は現時点では2種類、googleドライブにあったファイルを見て使うであろう物質を予め設定してありますが、他のものに変更はできます。
ただしnasa-ceaの方で対応してないと計算結果が出ません。N2Oのエネルギーを設定していないのは、元々nasa-ceaにN2Oのエネルギーのデータがあるからです。
ファイル名は任意で入力して下さい。（もしかしたら使えない文字があるかもしれない）

・inp出力
値を全て入力（N2Oのエネルギーは除く）した上でinp出力を押すとファイル名.inpを作成します。
このファイルはnasa-ceaが計算をするために必要なファイルなのでこれがないと計算できません。
ここで生成されたファイルはNASA-CEAフォルダに入ってます。

・cea起動
cea起動を押すとFCEA2mというソフトが起動します。ファイル名が入力されている場合はソフトが勝手にファイル名を入力し計算が開始されます。ファイル名が入力されてない場合は手動で入れてください。
計算が成功すると、入力した名前.outというファイルが作成されます。ここに計算結果がずらーっと書いてあります。

・excel出力
Excel出力を押すと、ファイル名.outというファイルを探してその中にある計算結果をSimulatorフォルダの中に.xlsxで出力します。Excelで開けます。
現時点で出力するデータはCSTARとGAMMA、計算条件、SimulationTemplate.xlsxがあれば計算シートを生成します
SimulationTemplate.xlsxは、プログラムではシートをまるごとコピー→O/Fなど自動入力するところはセル指定で入力しているのでそれ以外なら自由に変えても問題ありません

・個別計算
主に終了時O/Fと終了時燃焼室圧力だけを指定して計算した場合に使いそうな機能
O/F、燃焼室圧力がそれぞれ1つだけなら計算結果が1通りなのでその結果を表示します
