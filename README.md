# boiled-egg AutoHotkey 設定

`boiled-egg.ahk` は、ローマ字のIME再入力、IME状態の制御、Emacs風キー操作、プレーンテキスト貼り付けをまとめた AutoHotkey v2 スクリプトです。

## 必要ファイル

同じフォルダーに次の2ファイルを置きます。

```text
boiled-egg.ahk
IMEv2.ahk
```

メインスクリプトは次の記述で `IMEv2.ahk` を読み込みます。

```ahk
#Include IMEv2.ahk
```

### IMEv2.ahkの開発元と入手方法

`IMEv2.ahk` は、eamatさんが作成したAutoHotkey v1系のIME制御ライブラリ `IME.ahk` を基に、綾木 健一郎さんがAutoHotkey v2.0向けへ移植したIME制御関数群です。

* 開発・配布リポジトリ: [k-ayaki/IMEv2.ahk](https://github.com/k-ayaki/IMEv2.ahk)
* ソースファイル: [IMEv2.ahkを表示](https://github.com/k-ayaki/IMEv2.ahk/blob/master/IMEv2.ahk)
* 解説記事: [AutoHotkey v2.0のIME制御用 関数群 IMEv2.ahk](https://qiita.com/kenichiro_ayaki/items/d55005df2787da725c6f)

入手手順:

1. [GitHubリポジトリ](https://github.com/k-ayaki/IMEv2.ahk)を開きます。
2. リポジトリをZIPで取得するか、`IMEv2.ahk` のソース画面からファイルを保存します。
3. 保存したファイル名が `IMEv2.ahk` であることを確認します。
4. `boiled-egg.ahk` と同じフォルダーへ配置します。

Gitを使用する場合は、次のコマンドでも取得できます。

```powershell
git clone https://github.com/k-ayaki/IMEv2.ahk.git
```

このREADMEの構成では、リポジトリ内の `IMEv2.ahk` だけをメインスクリプトと同じフォルダーへコピーします。

## 動作環境

* Windows
* AutoHotkey v2
* 日本語IME
* UI Automation対応アプリ、またはコピー操作で選択文字列を取得できるアプリ

## 起動方法

1. AutoHotkey v2をインストールします。
2. `boiled-egg.ahk` と `IMEv2.ahk` を同じフォルダーへ置きます。
3. `boiled-egg.ahk` を実行します。
4. 古い版が起動中の場合は、先に終了してください。

`#SingleInstance Force` を使用しているため、同じスクリプトを再実行すると旧インスタンスは置き換えられます。

## ホットキー一覧

|ホットキー|動作|
|-|-|
|`Ctrl+J`|キャレット直前のローマ字列を選択してIMEへ再入力|
|`Ctrl+H`|Backspace|
|`Ctrl+M`|IME OFF時はEnter、IME ON時はIMEをOFF|
|`Ctrl+O`|IME ON/OFFを切り替え|
|`無変換`|IME OFF|
|`変換`|IME ON、変換モードを設定|
|`Ctrl+Win+Alt+V`|プレーンテキストとして貼り付け|
|`Ctrl+Esc`|スクリプト終了|

AutoHotkeyの記号表現では、`^` がCtrl、`#` がWin、`!` がAltです。

## Ctrl+Jによる再入力

### 基本動作

キャレットの直前から左方向へ文字を調べ、対象の半角文字列を選択して日本語IMEへ再入力します。

例:

```text
soukamo
```

末尾で `Ctrl+J` を押すと、IMEへ `soukamo` を再入力します。

```text
そうかも
```

### 対象文字

左スキャンの対象は次の半角ASCII文字だけです。

```text
A-Z
a-z
0-9
. , ! ? -
```

### 全角文字の扱い

次のような非ASCII文字を検出すると、その位置で左スキャンを終了します。

* ひらがな、カタカナ、漢字
* 全角英数字
* 全角記号
* 全角の `？` や `！`
* その他のUnicode文字

例:

```text
？Soudesuka?
```

右端で `Ctrl+J` を押すと、左側の全角 `？` は境界として残し、半角部分だけを再入力します。

```text
？そうですか？
```

IMEへ送るローマ字は小文字化されます。そのため `Soudesuka?` は `soudesuka?` として送信されます。

## 対応経路

### メモ帳

UI AutomationのTextPatternを使用して、選択中の文字列を取得します。

メモ帳のUIAが選択範囲の前後へCR/LFを付加する場合は、比較時だけ前後の改行を除外します。文字列内部の改行は変更しません。

### Excel

Excelでは次の3状態に対応します。

#### 1\. セルを選択した状態

* Excelの編集ウィンドウがないことを確認
* F2を送信
* セル末尾の編集モードへ移行
* クリップボード経路で文字列を取得
* 対象文字列をDeleteで削除
* IMEをONにしてローマ字を再入力

#### 2\. セル内にキャレットがある状態

* `EXCEL6` クラスの子ウィンドウを検出
* すでに編集モードと判断
* F2を追加送信しない
* 現在のキャレット位置から左スキャン

#### 3\. 数式バーを編集中

* UI AutomationまたはExcel編集状態を確認
* F2を追加送信しない
* 数式バー内のキャレット位置から左スキャン

Excelのクリップボード結果にタブやCR/LFが含まれる場合は、セル範囲コピーとして扱い、不正な変換を防ぎます。

### Emacsなど

UI Automationで選択文字列を取得できない場合は、`Ctrl+Insert` によるコピーへフォールバックします。

コピー後に選択が解除されるアプリでは、1文字ずつ左へ走査した後、対象範囲を選び直して処理します。

## IME制御

スクリプト起動時はIMEをOFFにします。

アプリが意図せずIMEをONにした場合にOFFへ戻す監視タイマーがあります。ただし、`Ctrl+J` の再入力処理中は競合を避けるため、タイマー自体を一時停止します。

IME再入力時は次を行います。

1. IMEをON
2. IME ON状態を確認
3. 変換モードを25へ設定
4. `SendEvent` でローマ字をキー入力

## Ctrl+Win+Alt+Vによるプレーンテキスト貼り付け

書式付きテキストをプレーンテキストとして貼り付けます。

処理手順:

1. 元のクリップボード全体を `ClipboardAll()` で保存
2. `A\_Clipboard` からテキストだけを取得
3. クリップボードをテキスト形式だけに一時変更
4. `Ctrl+V` を送信
5. 元のクリップボードを復元

画像だけがコピーされているなど、テキスト形式を取得できない場合は貼り付けません。

プレーンテキスト貼り付け後も、元のHTML、画像、書式付きデータはクリップボードへ戻されます。

## 待機時間とタイムアウト

スクリプト先頭の定数で調整できます。

```ahk
EXCEL\_F2\_TIMEOUT\_MS := 300
EXCEL\_DELETE\_SETTLE\_MS := 30
EXCEL\_IME\_SETTLE\_MS := 25
CLIPBOARD\_TIMEOUT\_MS := 300
UIA\_BOUNDARY\_TIMEOUT\_MS := 180
EXCEL\_BOUNDARY\_TIMEOUT\_MS := 300
```

|設定|用途|
|-|-|
|`EXCEL\_F2\_TIMEOUT\_MS`|F2後にExcel編集モードを確認する最大時間|
|`EXCEL\_DELETE\_SETTLE\_MS`|Excelで対象文字列を削除した後の待機|
|`EXCEL\_IME\_SETTLE\_MS`|ExcelでIME設定後に再入力まで待つ時間|
|`CLIPBOARD\_TIMEOUT\_MS`|選択文字列のコピー完了待ち|
|`UIA\_BOUNDARY\_TIMEOUT\_MS`|UIA経路で境界文字を除外したことを確認する時間|
|`EXCEL\_BOUNDARY\_TIMEOUT\_MS`|Excelのクリップボード経路で境界を確認する時間|

Excelの再入力では、次のキー間隔を使用します。

```ahk
SetKeyDelay(25, 10)
```

その他の経路では次を使用します。

```ahk
SetKeyDelay(35, 25)
```

## ログ

スクリプトと同じフォルダーへ次のログを出力します。

```text
boiled-egg-debug.log
```

主なログ項目:

|項目|内容|
|-|-|
|`start`|処理開始、対象プロセス、ウィンドウID|
|`excel-state`|Excelのセル選択、セル内編集、数式バーなどの判定結果|
|`scan`|左スキャンで取得した文字列と方式|
|`boundary-detected`|全角文字などの境界を検出|
|`boundary-verify`|境界文字が選択範囲から外れたことを確認|
|`ime-on-ok`|IME ONを確認|
|`reinput`|IMEへ再入力する文字列|
|`success`|正常終了と総所要時間|
|`exception`|エラー内容と行番号|

成功例:

```text
success | text="soukamo" totalElapsed=...
```

## トラブルシューティング

### Ctrl+Jでセルが移動する

Excelで編集状態を誤判定している可能性があります。ログの `excel-state` を確認します。

セル内編集中の期待値:

```text
mode=already-editing-native
```

セル選択状態からF2で入った場合の期待値:

```text
mode=entered-editing-native
```

### `TextPattern unavailable` と記録される

UI Automationを提供しない編集欄では、クリップボード経路へ自動的にフォールバックします。後続の `scan` と `success` があれば正常です。

### 変換されずローマ字のままになる

次のログを確認します。

```text
ime-set-called | allowed=1 state=1
ime-on-ok | state=1
conv-mode | mode=25
reinput
success
```

### 全角文字まで選択される

最新版ではコードポイントが `0x7F` を超える文字を境界にします。旧版ではなく、`boiled-egg.ahk` を起動していることを確認してください。

### プレーンテキスト貼り付けが動かない

* Ctrl、Win、Altをすべて押しながらVを押す
* クリップボードにテキストがあることを確認
* 他の常駐ソフトが同じホットキーを使用していないか確認

## 終了

次のどちらかで終了できます。

```text
Ctrl+Esc
```

または、タスクトレイのAutoHotkeyアイコンから終了します。

