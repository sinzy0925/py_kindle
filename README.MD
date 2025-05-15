# README.MD
これはcodeB.py を使ってもらうためのものです。

codeB.pyを自分が一元化したいソースコードのトップディレクトリにおいて、実行してください。
深いディレクトリにあるファイルも一元化してくれます。
実行の前に、モジュールのインストールが必要です。
pip install pathspec

次に、
python codeB.py を実行すると同じディレクトリにcode_output.txtが生成されます。

codeB.pyをソースを一元化したいフォルダへコピーして、実行すればそのアプリのソースが一元化されます。


このファイル構成ですと、以下のようになってます。
同じディレクトリにあっても一元化したいもの、したく無いものがあると思うので
codeB.pyで設定できます
今回の例では、敢えて同じディレクトリのcodeB.py、code_output.txt、kindle_shot2.pyは除外する設定です。
それぞれのファイル名を記載してますので、 code_output.txtを見ると分かりやすいと思います。

除外したいファイルを書く場所は、
codeB.pyの、
HARDCODED_EXCLUDE_PATTERNSの配列にファイル名やディレクトリ名を入れれば良いです。

ディレクトリ構成でいうとこうなります。
C:\Users\sinzy\py_kindle
│  codeB.py         <=除外対象
│  code_output.txt  <=除外対象
│  kindle_shot.py   <=出力
│  kindle_shot2.py  <=除外対象
│  README.MD        <=出力
│
└─test1
        kindle_shot3.py <=除外対象（test1フォルダを除外にしている）
└─test2
        kindle_shot4.py <=出力

以上です。