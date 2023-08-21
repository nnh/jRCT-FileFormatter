# jRCT ファイル整形プログラム ユーザーガイド

このガイドでは、jRCT ファイル整形プログラムの使用方法と事前準備について説明します。このプログラムは、変更前と変更後の jRCT ファイルを比較し、適切な形式に整形するためのツールです。

## 事前準備

1. プログラムを使用するためには、jRCT の規定に準じた形式の `.xlsx` ファイルが必要です
1. ファイルは以下のディレクトリに格納してください。変更前後でファイル名が同じでも問題なく動作します。
   - `input/before` フォルダ: 変更前のファイルをここに格納します。
   - `input/after` フォルダ: 変更後のファイルをここに格納します。

### 制限事項

- 各フォルダには `.xlsx` ファイルが 1 つだけ格納されている必要があります。
- `.xlsx` ファイルが存在しない場合、または 2 つ以上の `.xlsx` ファイルが格納されている場合、処理は中断されます。
- `.xlsx` ファイル以外（テキストや PDF など）が入っていても処理に影響しません。

## プログラム実行

1. `programs/jRCT-FileFormatter.xlsm` ファイルを開いてください。
2. 事前に準備したファイルを格納した後、プログラムを実行したい場合は「実行」ボタンをクリックしてください。

## 出力

プログラムを実行すると、`output` フォルダにファイルが作成されます。作成されたファイルは以下の手順で整形されます:

- `input/before` フォルダ内のファイル（`bef.xlsx` とする）と `input/after` フォルダ内のファイル（`aft.xlsx` とする）を比較します。
- 比較のためのキーとして、L 列「所属部署の郵便番号」と O 列「所属機関の住所」を使用します。
- `bef.xlsx` に存在し、`aft.xlsx` に存在しない行は「削除」の処理を行います。
- `bef.xlsx` と `aft.xlsx` の両方に存在する行はそのまま出力します。
- `bef.xlsx` に存在しないが `aft.xlsx` に存在する行は最後の行に追加されます。

## 削除の処理

削除の処理は以下のように行われます:

- A, B, I, K, N, Q, R, S, T, W, AA, AB, AE 列に「削除」の文字列が出力されます。
- E, G, J 列に「X」の文字列が出力されます。
- F, H 列に空白が出力されます。
- L, U 列に「000-0000」の文字列が出力されます。
- M, V 列に「その他」の文字列が出力されます。
- O 列に「00-0000-0000」の文字列が出力されます。
- P, Z 列に「X@X.com」の文字列が出力されます。
- X, Y 列に「000-000-0000」の文字列が出力されます。
- AC 列に「無効」の文字列が出力されます。
- AD 列はそのまま出力されます。
