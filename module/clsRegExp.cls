VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/**
' * @dependency Microsoft VBScript Regular Expressions
' * @dependency Microsoft Scripting Runtime
' */

Private objRegExp As RegExp
Private objFSO As FileSystemObject

Private propCode As Long
Private propMessage As String

Private propPath As String
Private propDestination As String
Private propPattern As String
Private propReplacement As String
Private propGlobalMatch As Boolean

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
' * @param strDestination ファイルのパス。
' */
Public Property Let destination(strDestination As String)
    propDestination = strDestination
End Property

'/**
' * @return 出力先となるファイルのパス。
' */
Public Property Get destination() As String
    destination = propDestination
End Property

'/**
' * 操作対象とする文字列の検索パターンを正規表現で指定する。
' *
' * @param strPattern パターン文字列。
' */
Public Property Let pattern(strPattern As String)
    propPattern = strPattern
End Property

'/**
' * @return パターン文字列。
' */
Public Property Get pattern() As String
    pattern = propPattern
End Property

'/**
' * 置換後の文字列、または挿入する文字列を指定する。
' *
' * @param strReplacement 置換後の文字列、または挿入する文字列。
' */
Public Property Let replacement(strReplacement As String)
    propReplacement = strReplacement
End Property

'/**
' * @return 置換後の文字列、または挿入する文字列。
' */
Public Property Get replacement() As String
    replacement = propReplacement
End Property

'/**
' * 文字列全体を操作対象とするかどうかを指定する。
' *
' * @param boolGlobal
' *     False: 検索パターンに最初に合致した文字列のみを操作対象とする。(既定値)
' *     True: 検索パターンに合致した文字列全てを操作対象とする。
' */
Public Property Let globalMatch(boolGlobal As Boolean)
    propGlobalMatch = boolGlobal
End Property

'/**
' * @return True / False
' */
Public Property Get globalMatch() As Boolean
    globalMatch = propGlobalMatch
End Property

'/**
' * 初期化処理。
' */
Private Sub Class_Initialize()
    Set objRegExp = New RegExp
    Set objFSO = New FileSystemObject
    propGlobalMatch = False
End Sub

'/**
' * 検索パターンに合致する文字列を持つ行を新しいファイルへ書き出す。
' */
Public Sub selectString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.selectString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.selectString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        
        If objRegExp.Test(strLine) Then
            Print #lngDstNumber, strLine
        End If
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.selectString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.selectString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub

'/**
' * 検索パターンに合致する文字列を持つ行を削除する。
' */
Public Sub deleteString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.deleteString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.deleteString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        
        If objRegExp.Test(strLine) = False Then
            Print #lngDstNumber, strLine
        End If
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.deleteString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.deleteString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub

'/**
' * 検索パターンに合致する部分文字列を置換して新しいファイルへ書き出す。
' *
' * @param strReplacement 置換後の文字列を指定する。
' */
Public Sub replaceSubString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.replaceSubString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.replaceSubString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        Print #lngDstNumber, objRegExp.Replace(strLine, replacement)
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.replaceSubString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.replaceSubString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub

'/**
' * 検索パターンに合致する文字列を持つ行全体を新しい文字列に置換する。
' */
Public Sub replaceString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.replaceString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.replaceString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        
        If objRegExp.Test(strLine) Then
            Print #lngDstNumber, replacement
        Else
            Print #lngDstNumber, strLine
        End If
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.replaceString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.replaceString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub

'/**
' * 検索パターンに合致する文字列を持つ行の直前に新しい行を挿入する。
' */
Public Sub insertString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.insertString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.insertString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        
        If objRegExp.Test(strLine) Then
            Print #lngDstNumber, replacement & vbCrLf & strLine
        Else
            Print #lngDstNumber, strLine
        End If
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.insertString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.insertString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub

'/**
' * 検索パターンに合致する文字列を持つ行の直後に新しい行を挿入する。
' */
Public Sub addString()
    On Error GoTo errExit2
    
    Dim strLine As String
    Dim lngSrcNumber As Long
    Dim lngDstNumber As Long
    
    With objRegExp
        .pattern = pattern
        .Global = globalMatch
    End With
    
    With objFSO
        If .FileExists(path) = False Then
            propMessage = "clsRegExp.addString path の値が不正です。"
            GoTo errExit1
        End If
        
        If .FolderExists(.GetParentFolderName(destination)) = False Then
            propMessage = "clsRegExp.addString destination の値が不正です。"
            GoTo errExit1
        End If
    End With
    
    lngSrcNumber = FreeFile
    Open path For Input As #lngSrcNumber
    
    lngDstNumber = FreeFile
    Open destination For Output As #lngDstNumber
    
    Do Until EOF(lngSrcNumber)
        Line Input #lngSrcNumber, strLine
        
        If objRegExp.Test(strLine) Then
            Print #lngDstNumber, strLine & vbCrLf & replacement
        Else
            Print #lngDstNumber, strLine
        End If
    Loop
    
    Close lngSrcNumber
    Close lngDstNumber
    
    propCode = 2
    propMessage = "clsRegExp.addString [" & destination & "] 完了しました。"
    Debug.Print message
    Exit Sub
    
errExit1:
    MsgBox message
    propCode = 1
    Exit Sub

errExit2:
    propMessage = "clsRegExp.addString エラーが発生しました。" _
        & " コード: " & Err.Number _
        & " メッセージ: " & Err.Description
    MsgBox message
    propCode = 1
    Close lngSrcNumber
    Close lngDstNumber
End Sub
