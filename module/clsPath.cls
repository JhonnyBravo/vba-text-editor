VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/**
' * @dependency Microsoft Scripting Runtime
' */

Private objFSO As FileSystemObject

Private propCode As Long
Private propMessage As String

Private propPath As String

'/**
' * コマンド終了直後の終了コードを返す。
' *
' * @return
' *     0: 変更も異常もなく終了したことを表す。
' *     1: 異常終了したことを表す。
' *     2: 操作対象に対して何らかの変更を加えて正常終了したことを表す。
' */
Public Property Get code() As Long
    code = propCode
End Property

'/**
' * @return コマンド終了直後の終了メッセージを返す。
' */
Public Property Get message() As String
    message = propMessage
End Property

'/**
' * 操作対象とするファイルまたはフォルダーのパスを指定する。
' *
' * @param strPath ファイルまたはフォルダーのパス。
' */
Public Property Let path(strPath As String)
    propPath = strPath
End Property

'/**
' * @return ファイルまたはフォルダーのパス。
' */
Public Property Get path() As String
    path = propPath
End Property

'/**
' * 初期化処理。
' */
Private Sub Class_Initialize()
    Set objFSO = New FileSystemObject
End Sub

'/**
' * ファイルまたはフォルダーのパスから末尾の部分のみを取得して返す。
' */
Public Function getBaseName() As String
    On Error GoTo errExit
    
    Dim result As String
    
    result = objFSO.GetFileName(path)
    propCode = 2
    propMessage = "clsPath.getBaseName [" & result & "] パスを取得しました。"
    
    getBaseName = result
    Debug.Print message
    Exit Function

errExit:
    propMessage = "clsPath.getBaseName エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Function

'/**
' * ファイルまたはフォルダーのパスから親フォルダーのパスを取得して返す。
' */
Public Function getDirectoryName() As String
    On Error GoTo errExit
    
    Dim result As String
    
    result = objFSO.GetParentFolderName(path)
    propCode = 2
    propMessage = "clsPath.getDirectoryName [" & result & "] パスを取得しました。"
    
    getDirectoryName = result
    Debug.Print message
    Exit Function

errExit:
    propMessage = "clsPath.getDirectoryName エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Function
