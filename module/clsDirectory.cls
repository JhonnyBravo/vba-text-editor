VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirectory"
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
' * 操作対象とするフォルダーのパスを指定する。
' *
' * @param strPath フォルダーのパス。
' */
Public Property Let path(strPath As String)
    propPath = strPath
End Property

'/**
' * @return フォルダーのパス。
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
' * フォルダーを作成する。
' */
Public Sub create()
    On Error GoTo errExit
    
    With objFSO
        If .FolderExists(.GetParentFolderName(path)) = False Then
            propCode = 1
            propMessage = "clsDirectory.create path の値が不正です。"
            MsgBox message
            Exit Sub
        End If
        
        If .FolderExists(path) Then
            propCode = 0
            propMessage = ""
        Else
            .CreateFolder path
            propCode = 2
            propMessage = "clsDirectory.create [" & path & "] フォルダーを作成しました。"
            Debug.Print message
        End If
    End With
    Exit Sub

errExit:
    propMessage = "clsFolder.create エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Sub

'/**
' * フォルダーを削除する。
' */
Public Sub delete()
    On Error GoTo errExit
    
    With objFSO
        If .FolderExists(path) = False Then
            propCode = 0
            propMessage = ""
        Else
            .DeleteFolder path
            propCode = 2
            propMessage = "clsDirectory.delete [" & path & "] フォルダーを削除しました。"
            Debug.Print message
        End If
    End With
    Exit Sub

errExit:
    propMessage = "clsDirectory.delete エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
End Sub
