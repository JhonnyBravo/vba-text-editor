# CopyFile
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: 異常もリソースの変更もなく終了した状態を表す。
* 1: 異常終了した状態を表す。
* 2: リソースの変更に成功した状態を表す。

ファイルをコピーする。

## Method
### runCommand

```
Public Sub runCommand(
  strSrcPath As String, 
  strDstPath As String
)
```

**パラメータ:**

* strSrcPath - コピー元ファイルのパスを指定する。
* strDstPath - コピー先のパスを指定する。

[HOME](index)
