# GetFilePath
## Property
### code

```
Public Property Get code() As Long
```

メソッド実行直後の終了コードを返す。

* 0: エラーもなく、リソースの変更もなく終了した状態を表す。
* 1: エラー終了した状態を表す。
* 2: リソースを変更し、正常終了した状態を表す。

### filterCollection

```
Public Property Get filterCollection() As Collection
```

ファイルピッカーの拡張子フィルターへ渡す配列を格納するコレクションを返す。  
配列は下記の形式で登録する。

```
Array("Images", "*.gif; *.jpg; *.jpeg")
```

## Method
### runCommand

```
Public Function runCommand() As String
```

ファイルピッカーを起動し、選択したファイルのパスを取得して返す。

**戻り値:**

選択したファイルのパスを返す。

[HOME](index)
