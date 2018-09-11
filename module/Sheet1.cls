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
        
        Dim objCD As New CreateDirectory
        
        With objCD
            .runCommand strDstPath
            
            If .code = 1 Then
                Exit Sub
            End If
        End With
        
        Dim objGBN As New GetBaseName
        Dim strBaseName As String
        
        With objGBN
            strBaseName = .runCommand(strSrcPath)
            
            If .code = 1 Then
                Exit Sub
            End If
        End With

        Dim objFC As New FileController
        
        With objFC
            .copy strSrcPath, strDstPath & "\" & strBaseName
            
            If .code = 1 Then
                Exit Sub
            End If
        End With
        
        strRecordEnd = .Range("B6").End(xlDown).Address(False, False)
        Set objRange = .Range(strRecordEnd, "E6")
        varActions = objRange
        lngRowMax = UBound(varActions)
        
        Dim objREC As New RegExpController
        Dim strAction As String
        Dim strPattern As String
        Dim strReplacement As String
        Dim strAddPosition As String
        
        For lngRow = 1 To lngRowMax
            strAction = varActions(lngRow, 1)
            strPattern = varActions(lngRow, 2)
            strReplacement = varActions(lngRow, 3)
            
            With objFC
                .create strDstPath & "\tmp.txt"
                
                If .code = 1 Then
                    Exit Sub
                End If
            End With
            
            Select Case LCase(strAction)
                Case "grep"
                    With objREC
                        .selectString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "delete"
                    With objREC
                        .deleteString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "replace"
                    With objREC
                        .replaceSubString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern, strReplacement
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "change"
                    With objREC
                        .replaceString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern, strReplacement
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "insert"
                    With objREC
                        .insertString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern, strReplacement
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                Case "add"
                    With objREC
                        .addString strDstPath & "\" & strBaseName, strDstPath & "\tmp.txt", strPattern, strReplacement
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
                    
                    With objFC
                        .copy strDstPath & "\tmp.txt", strDstPath & "\" & strBaseName
                        
                        If .code = 1 Then
                            Exit Sub
                        End If
                    End With
            End Select
            
            With objFC
                .delete strDstPath & "\tmp.txt"
                
                If .code = 1 Then
                    Exit Sub
                End If
            End With
        Next
    End With
    
    MsgBox "完了しました。"
End Sub

Private Sub btnGetPath_Click()
    Dim objGFP As New GetFilePath
    Dim strPath As String
    
    With objGFP
        strPath = .runCommand
        
        If .code = 2 Then
            Sheets("Main").Range("C2").Value = strPath
        End If
    End With
End Sub
