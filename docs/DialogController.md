# DialogController
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: エラーもなく、リソースの変更もなく終了した状態を表す。
* 1: エラー終了した状態を表す。
* 2: リソースを変更して正常終了した状態を表す。

## Method
### getCsvPath

```
Public Function getCsvPath() As String
```

ファイルピッカーを起動し、選択した CSV ファイルのパスを取得して返す。

**戻り値:**

path - 選択したファイルのパスを返す。

### getExcelPath

```
Public Function getExcelPath() As String
```

ファイルピッカーを起動し、選択した Excel ファイルのパスを取得して返す。

**戻り値:**

選択したファイルのパスを返す。

### getDirectoryPath

```
Public Function GetDirectoryPath() As String
```

ディレクトリピッカーを起動し、選択したディレクトリのパスを取得して返す。

**戻り値:**

選択したディレクトリのパスを返す。

[HOME](index)
