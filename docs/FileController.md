# FileController
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
### create

```
Public Sub create(strPath As String)
```

ファイルを新規作成する。

**パラメータ:**

* strPath - 新規作成するファイルのパスを指定する。

### delete

```
Public Sub delete(strPath As String)
```

ファイルを削除する。

**パラメータ:**

* strPath - 削除対象とするファイルのパスを指定する。

### copy

```
Public Sub copy(
  strSrcPath As String,
  strDstPath As String
)
```

ファイルをコピーする。

**パラメータ:**

* strSrcPath - コピー元ファイルのパスを指定する。
* strDstPath - コピー先のパスを指定する。

[HOME](index)
