# GetBaseName
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: 異常もリソースの変更もなく終了した状態を表す。
* 1: 異常終了した状態を表す。
* 2: リソースの変更に成功した状態を表す。

## Method
### runCommand

```
Public Function runCommand(strPath As String) As String
```

指定したパスから、ファイル名またはディレクトリ名を取得する。

**パラメータ:**

* strPath - ファイルまたはディレクトリのパスを指定する。

**戻り値:**

ファイル名またはディレクトリ名を返す。

[HOME](index)
