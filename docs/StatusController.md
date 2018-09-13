# StatusController
## Property
### code

```
Public Property
  Get code() As Long
  Let code(lngCode As Long)
```

メソッド終了直後の終了コードを返す。

* 0: 異常もリソースの変更もなく終了した状態を表す。
* 1: 異常終了した状態を表す。
* 2: リソースの変更に成功した状態を表す。

**パラメータ:**

* lngCode - メソッド終了直後の終了コードを設定する。

### message

```
Public Property
  Get message() As String
  Let message(strMessage As String)
```

エラー発生時に出力するメッセージを返す。

**パラメータ:**

* strMessage - エラー発生時に出力するメッセージを設定する。

## Method
### initStatus

```
Public Sub initStatus()
```

終了コードとエラーメッセージを初期化する。

### errorTerminate

```
Public Sub errorTerminate()
```

終了コードを 1 に設定し、エラーメッセージを出力する。

[HOME](index)
