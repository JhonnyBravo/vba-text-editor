# GetTextStream
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: エラーもなく、リソースの変更もなく終了した状態を表す。
* 1: エラー終了した状態を表す。
* 2: リソースを変更し、正常終了した状態を表す。

## Method
### runCommand

```
Public Function runCommand(
  strPath As String,
  strMode As String
) As TextStream
```

**パラメータ:**
* strPath - TextStream を取得するファイルのパスを指定する。
* strMode -
  * get: ファイルを読み取りモードで開く。
  * set: ファイルを上書きモードで開く。
  * add: ファイルを追記モードで開く。

**戻り値:**

TextStream を取得して返す。

[HOME](index)
