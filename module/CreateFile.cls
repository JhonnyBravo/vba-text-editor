VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objSC As New StatusController
Private propCode As Long

'@return code メソッド実行直後の終了コードを返す。
'   0: 異常もリソースの変更もなく終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: リソースの変更に成功した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'ファイルを新規作成する。
'
'@param strPath 新規作成するファイルのパスを指定する。
Public Sub runCommand(strPath As String)
    objSC.initStatus
    
    On Error GoTo errExit
    
    Dim objFSO As New FileSystemObject
    
    With objFSO
        If .FileExists(strPath) Then
            With objSC
                .initStatus
                propCode = .code
            End With
        Else
            .CreateTextFile strPath
            
            With objSC
                .code = 2
                .message = ""
                propCode = .code
            End With
        End If
    End With
    
    Exit Sub
    
errExit:
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        propCode = .code
    End With
End Sub
