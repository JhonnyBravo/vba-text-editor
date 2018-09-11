VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SelectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objSC As New StatusController
Private propCode As Long

Private propGlobalMatch As Boolean

'@return code メソッド実行直後の終了コードを返す。
'   0: 異常もリソースの変更もなく終了した状態を表す。
'   1: 異常終了した状態を表す。
'   2: リソースの変更に成功した状態を表す。
Public Property Get code() As Long
    code = propCode
End Property

'@return GlobalMatch 正規表現による文字列操作の有効範囲を表す真偽値を返す。
Public Property Get GlobalMatch() As Boolean
    GlobalMatch = propGlobalMatch
End Property

'@param boolGlobalMatch
'   True: 正規表現パターンに合致する全ての文字列を操作対象とする。
'   False: 正規表現パターンに合致する最初の文字列のみを操作対象とする。
Public Property Let GlobalMatch(boolGlobalMatch As Boolean)
    propGlobalMatch = boolGlobalMatch
End Property

Private Sub Class_Initialize()
    propGlobalMatch = False
End Sub

'正規表現パターンに合致する文字列を持つ行のみを抽出する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 取得対象とする文字列の正規表現パターンを指定する。
Public Sub runCommand(strSrcPath As String, strDstPath As String, strPattern As String)
    objSC.initStatus
    
    Dim objGTS As New GetTextStream
    Dim objSrcTS As TextStream
    Dim objDstTS As TextStream
    
    With objGTS
        Set objSrcTS = .runCommand(strSrcPath, "get")
        
        If .code = 1 Then
            Exit Sub
        End If
        
        Set objDstTS = .runCommand(strDstPath, "add")
        
        If .code = 1 Then
            Exit Sub
        End If
    End With
    
    On Error GoTo errExit
    
    Dim objRE As New RegExp
    
    With objRE
        .pattern = strPattern
        .Global = GlobalMatch
    End With
    
    Dim strLine As String
    Dim lngLineNumber As Long
    
    With objSrcTS
        Do Until .AtEndOfStream
            strLine = .ReadLine
            lngLineNumber = .Line
            
            If objRE.Test(strLine) Then
                objDstTS.WriteLine strLine
            End If
        Loop
    End With
    
    objSrcTS.Close
    objDstTS.Close
    
    With objSC
        .code = 2
        .message = ""
        propCode = .code
    End With
    
    Exit Sub
    
errExit:
    objSrcTS.Close
    objDstTS.Close
    
    With objSC
        .message = "エラーが発生しました。 " & Err.Description
        .errorTerminate
        Debug.Print lngLineNumber & " 行目: " & strLine
        propCode = .code
    End With
End Sub
