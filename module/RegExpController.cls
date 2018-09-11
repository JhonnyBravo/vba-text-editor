VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExpController"
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
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
Public Sub selectString(strSrcPath As String, strDstPath As String, strPattern As String)
    Dim objSS As New selectString
    
    With objSS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern
        propCode = .code
    End With
End Sub

'正規表現パターンに合致する文字列を持つ行を削除する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
Public Sub deleteString(strSrcPath As String, strDstPath As String, strPattern As String)
    Dim objDS As New deleteString
    
    With objDS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern
        propCode = .code
    End With
End Sub

'正規表現パターンに合致する文字列を新しい文字列に置換する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
'@param strReplacement 置換後の文字列を指定する。
Public Sub replaceSubString(strSrcPath As String, strDstPath As String, strPattern As String, strReplacement As String)
    Dim objRSS As New replaceSubString
    
    With objRSS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern, strReplacement
        propCode = .code
    End With
End Sub

'正規表現パターンに合致する文字列を持つ行全体を新しい文字列に置換する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
'@param strReplacement 置換後の文字列を指定する。
Public Sub replaceString(strSrcPath As String, strDstPath As String, strPattern As String, strReplacement As String)
    Dim objRS As New replaceString
    
    With objRS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern, strReplacement
        propCode = .code
    End With
End Sub

'正規表現パターンに合致する文字列を持つ行の一つ前の行に新しい文字列を挿入する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
'@param strNewLine 新規挿入する文字列を指定する。
Public Sub insertString(strSrcPath As String, strDstPath As String, strPattern As String, strNewLine As String)
    Dim objIS As New insertString
    
    With objIS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern, strNewLine
        propCode = .code
    End With
End Sub

'正規表現パターンに合致する文字列を持つ行の一つ後の行に新しい文字列を挿入する。
'
'@param strSrcPath 編集対象とするテキストファイルのパスを指定する。
'@param strDstPath 編集後の出力先として使用するファイルのパスを指定する。
'@param strPattern 対象とする文字列の正規表現パターンを指定する。
'@param strNewLine 新規挿入する文字列を指定する。
Public Sub addString(strSrcPath As String, strDstPath As String, strPattern As String, strNewLine As String)
    Dim objAS As New addString
    
    With objAS
        .GlobalMatch = GlobalMatch
        .runCommand strSrcPath, strDstPath, strPattern, strNewLine
        propCode = .code
    End With
End Sub
