# GetDirectoryPath
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
Public Function runCommand() As String
```

ファイルピッカーを起動し、選択したディレクトリのパスを取得して返す。

**戻り値:**

選択したディレクトリのパスを返す。

[HOME](index)
