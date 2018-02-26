VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFile"
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
Private propDestination As String

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
' * 操作対象とするファイルのパスを指定する。
' *
' * @param strPath ファイルのパス。
' */
Public Property Let path(strPath As String)
    propPath = strPath
End Property

'/**
' * @return 操作対象とするファイルのパス。
' */
Public Property Get path() As String
    path = propPath
End Property

'/**
' * 出力先となるファイルのパスを指定する。
' *
' * @param strPath ファイルのパス。
' */
Public Property Let destination(strPath As String)
    propDestination = strPath
End Property

'/**
' * @return 出力先となるファイルのパス。
' */
Public Property Get destination() As String
    destination = propDestination
End Property

'/**
' * 初期化処理。
' */
Private Sub Class_Initialize()
    Set objFSO = New FileSystemObject
End Sub

'/**
' * ファイルを作成する。
' */
Public Sub create()
    On Error GoTo errExit
    
    With objFSO
        If .FolderExists(.GetParentFolderName(path)) = False Then
            propCode = 1
            propMessage = "clsFile.create path の値が不正です。"
            MsgBox message
            Exit Sub
        End If
        
        If .FileExists(path) Then
            propCode = 0
            propMessage = ""
        Else
            .CreateTextFile path
            propCode = 2
            propMessage = "clsFile.create [" & path & "] ファイルを作成しました。"
            Debug.Print message
        End If
    End With
    Exit Sub

errExit:
    propMessage = "clsFile.create エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Sub

'/**
' * ファイルを削除する。
' */
Public Sub delete()
    On Error GoTo errExit
    
    With objFSO
        If .FileExists(path) = False Then
            propCode = 0
            propMessage = ""
        Else
            .DeleteFile path
            propCode = 2
            propMessage = "clsFile.delete [" & path & "] ファイルを削除しました。"
            Debug.Print message
        End If
    End With
    Exit Sub

errExit:
    propMessage = "clsFile.delete エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Sub

'/**
' * ファイルをコピーする。
' */
Public Sub copy()
    On Error GoTo errExit
    
    With objFSO
        If .FolderExists(.GetParentFolderName(path)) = False Then
            propCode = 1
            propMessage = "clsFile.create path の値が不正です。"
            MsgBox message
            Exit Sub
        ElseIf .FolderExists(.GetParentFolderName(destination)) = False Then
            propCode = 1
            propMessage = "clsFile.create destination の値が不正です。"
            MsgBox message
            Exit Sub
        End If
        
        .CopyFile path, destination, True
        propCode = 2
        propMessage = "clsFile.copy [" & destination & "] ファイルをコピーしました。"
        Debug.Print message
    End With
    Exit Sub

errExit:
    propMessage = "clsFile.copy エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Sub
