VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnEdit_Click()
    Dim strSrcPath As String
    Dim strRecordEnd As String
    Dim objRange As Range
    Dim varActions As Variant
    Dim lngRowMax As Long
    Dim lngRow As Long
    
    With Sheets("Main")
        strSrcPath = .Range("C2").Value
        
        If strSrcPath = "" Then
            MsgBox "Path を指定してください。"
            Exit Sub
        End If
        
        If .Range("B6").Value = "" Then
            MsgBox "Action を指定してください。"
            Exit Sub
        End If
        
        Dim strDirPath As String
        Dim strDstPath As String
        
        strDirPath = ActiveWorkbook.path
        strDstPath = strDirPath & "\output"
        
        Dim objDirectory As New clsDirectory
        Dim objFile As New clsFile
        
        With objDirectory
            .path = strDstPath
            .create
            
            If .code = 1 Then
                Exit Sub
            End If
        End With
        
        Dim objPath As New clsPath
        Dim strBasePath As String
        
        With objPath
            .path = strSrcPath
            strBasePath = .getBaseName
        End With
        
        With objFile
            .path = strSrcPath
            .destination = strDstPath & "\" & strBasePath
            .copy
            
            If .code = 1 Then
                Exit Sub
            End If
        End With

        strRecordEnd = .Range("B6").End(xlDown).Address(False, False)
        Set objRange = .Range(strRecordEnd, "E6")
        varActions = objRange
        lngRowMax = UBound(varActions)
        
        Dim objRegExp As New clsRegExp
        Dim strAction As String
        Dim strPattern As String
        Dim strReplacement As String
        Dim strAddPosition As String
        
        For lngRow = 1 To lngRowMax
            strAction = varActions(lngRow, 1)
            strPattern = varActions(lngRow, 2)
            strReplacement = varActions(lngRow, 3)
            
            Select Case LCase(strAction)
                Case "grep"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .selectString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "delete"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .deleteString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "replace"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .replacement = strReplacement
                        .replaceSubString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "change"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .replacement = strReplacement
                        .replaceString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "insert"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .replacement = strReplacement
                        .insertString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "add"
                    With objRegExp
                        .path = strDstPath & "\" & strBasePath
                        .destination = strDstPath & "\tmp.txt"
                        .pattern = strPattern
                        .replacement = strReplacement
                        .addString
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFile
                        .path = strDstPath & "\tmp.txt"
                        .destination = strDstPath & "\" & strBasePath
                        .copy
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                        
                        .delete
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
            End Select
        Next
    End With
    
    MsgBox "完了しました。"
End Sub

Private Sub btnGetPath_Click()
    Dim objPD As New clsPathDialog
    Dim strPath As String
    
    With objPD
        .pathType = "file"
        strPath = .getPath
        
        If .code = 2 Then
            Sheets("Main").Range("C2").Value = strPath
        End If
    End With
End Sub
