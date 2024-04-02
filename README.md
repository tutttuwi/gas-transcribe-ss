# gas-transcribe-ss
スプレッドシート転記ツール

## スプレッドシート

<https://docs.google.com/spreadsheets/d/15SRByQcZjw-cWdBKKNwqN6uPOGkFBmPLHIMmNDsGIxo/copy>

## できること

- スプレッドシート間のデータ転記が可能
- 転記前に初期化をすることもできます
- 複数のシートの情報を1つのシートにマージすることもできます

## 本ツール作成方法

ChatGPTへ以下の依頼後、コードを手直ししてます。

```
あなたはGoogleAppScriptが書けるエキスパートエンジニアです。

スプレッドシートの転記元情報から、転記先にデータを転記するための、
以下の処理をGoogleAppScriptで書いてください。

1. 「設定」シートから以下の情報を読み取る
　B4セル：転記元スプレッドシートID
　C4セル：転記元スプレッドシートリンク
　D4セル：転記元シート名
　E4セル：転記元セル範囲
　F4セル：転記先スプレッドシートID
　G4セル：転記先スプレッドシートリンク
　H4セル：転記先シート名
　I4セル：転記先セル範囲
　J4セル：初期化有無
　K4セル：追記有無
　　※5行目以降もデータがあれば複数配列レコードとして読み込む
　
2. 1.で読み取った情報をもとに以下の処理を実行する
　※複数レコードあれば、以下の処理を繰り返す
　
　・転記元スプレッドシートの情報をもとに、転記先スプレッドシートに転記する
　　※条件
　　　「初期化有無」の値がTRUEの場合、転記先シートの情報を全削除してから転記する
　　　「追記有無」の値がTRUEの場合、転記先シートの最終行から転記する

3. スプレッドシートの転記に成功した場合、
　「設定」シートのL4セルに現在日時を記入し、M4セルに「OK」と記入する
　 スプレッドシートの転記に失敗した場合、
　「設定」シートのL4セルに現在日時を記入し、M4セルに「NG」と記入する
　※1.のレコードが複数ある場合は、対応する行に上記の記入をしてください
　


```

以上🙆‍♂️



