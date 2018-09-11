VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileController"
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
Public Sub create(strPath As String)
    Dim objCF As New CreateFile
    
    With objCF
        .runCommand strPath
        propCode = .code
    End With
End Sub

'ファイルを削除する。
'
'@param strPath 削除対象とするファイルのパスを指定する。
Public Sub delete(strPath As String)
    Dim objDF As New DeleteFile
    
    With objDF
        .runCommand strPath
        propCode = .code
    End With
End Sub

'ファイルをコピーする。
'
'@param strSrcPath コピー元ファイルのパスを指定する。
'@param strDstPath コピー先のパスを指定する。
Public Sub copy(strSrcPath As String, strDstPath As String)
    Dim objCF As New CopyFile
    
    With objCF
        .runCommand strSrcPath, strDstPath
        propCode = .code
    End With
End Sub
