# ReplaceString
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: 異常もリソースの変更もなく終了した状態を表す。
* 1: 異常終了した状態を表す。
* 2: リソースの変更に成功した状態を表す。

### GlobalMatch

```
Public Property
  Get GlobalMatch() As Boolean
  Let GlobalMatch(boolGlobalMatch As Boolean)
```

正規表現による文字列操作の有効範囲を表す真偽値を返す。

* True: 正規表現パターンに合致する全ての文字列を操作対象とする。
* False: 正規表現パターンに合致する最初の文字列のみを操作対象とする。

**パラメータ:**

* boolGlobalMatch - 正規表現による文字列操作の有効範囲を表す真偽値を指定する。

## Method
### runCommand

```
Public Sub runCommand(
  strSrcPath As String,
  strDstPath As String,
  strPattern As String,
  strReplacement As String
)
```

正規表現パターンに合致する文字列を持つ行全体を新しい文字列に置換する。

**パラメータ:**

* strSrcPath - 編集対象とするテキストファイルのパスを指定する。
* strDstPath - 編集後の出力先として使用するファイルのパスを指定する。
* strPattern - 置換対象とする文字列の正規表現パターンを指定する。
* strReplacement - 置換後の文字列を指定する。

[HOME](index)
