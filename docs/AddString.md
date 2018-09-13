# AddString
## Property
### code

```
Public Property Get code() AS Long
```

メソッド実行直後の終了コードを返す。(読取専用)

* 0: 異常もリソースの変更もなく終了した状態を表す。
* 1: 異常終了した状態を表す。
* 2: リソースの変更に成功した状態を表す。

### GlobalMatch

```
Public Property
  Get GlobalMatch() AS Boolean
  Let GlobalMatch(boolGlobalMatch AS Boolean)
```

正規表現による文字列操作の有効範囲を表す真偽値を返す。

* True: 正規表現に合致するすべての文字列を操作対象とする。
* False: 正規表現に合致する最初の文字列のみを操作対象とする。

**パラメータ:**

* boolGlobalMatch - 正規表現による文字列操作の有効範囲を表す真偽値を指定する。

## Method
### runCommand

```
Public Sub runCommand(
  strSrcPath AS String,
  strDstPath AS String,
  strPattern AS String,
  strNewLine AS String
)
```

正規表現パターンに合致する文字列の一行後に新しい文字列を挿入する。

**パラメータ:**

* strSrcPath - 編集対象とするテキストファイルのパスを指定する。
* strDstPath - 編集後の出力先として使用するファイルのパスを指定する。
* strPattern - 検索対象とする文字列の正規表現パターンを指定する。
* strNewLine - 新規挿入する文字列を指定する。

[HOME](index)
