Private Sub Mandatory(i As Integer, j As Integer)
i = i + 11


        If (ActiveSheet.Cells(i, j).value = "") Then
            'Debug.Print j
            ActiveSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(i, 58).value = "Validation Failed"
        Else
            ' below line is commented by Sudipta
            ' ActiveSheet.Cells(i, j).Interior.Color = xlNone
            ' ActiveSheet.Cells(i, 58).value = "Validation Successful"
        End If
i = i - 11
End Sub
'Registry party ID

Sub ValidateAndFormatColumnC(rowid As Integer)


    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 3)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "4" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '4'."
            On Error GoTo 0
        End If
    End If

End Sub


'Registry site Id
Sub ValidateAndFormatColumnD(rowid As Integer)

    ' Declare variables
    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 4)
    value = targetCell.value
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    If targetCell.value <> "" Then

        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "2" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else

            targetCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"

            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '2'."
            On Error GoTo 0
        End If
    End If

End Sub


'SFDC Id
Sub ValidateAndFormatColumnAI(rowid As Integer)

    ' Declare variables
    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 35)
    value = targetCell.value
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    If targetCell.value <> "" Then

        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "2" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else

            targetCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"

            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '2'."
            On Error GoTo 0
        End If
    End If

End Sub
'Bill To Target Registry Party ID
Sub ValidateAndFormatColumnAT(rowid As Integer)

        Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 46)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "4" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '4'."
            On Error GoTo 0
        End If
    End If

End Sub
'Bill To Target Registry Account ID
Sub ValidateAndFormatColumnAU(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 47)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "5" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '5'."
            On Error GoTo 0
        End If
    End If
        
    

End Sub
'Ship To Target Registry Party ID
Sub ValidateAndFormatColumnAW(rowid As Integer)

        Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 49)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "4" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '4'."
            On Error GoTo 0
        End If
    End If

End Sub
'Ship To Target Registry Account ID
Sub ValidateAndFormatColumnAX(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 50)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "5" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '5'."
            On Error GoTo 0
        End If
    End If
        
End Sub
'Relationship Target Registry Party ID
Sub ValidateAndFormatColumnBA(rowid As Integer)

        Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 53)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "4" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '4'."
            On Error GoTo 0
        End If
    End If
End Sub
'Relationship Target Registry Account ID
Sub ValidateAndFormatColumnBB(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 54)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0
    
    If targetCell.value <> "" Then
        If IsNumeric(value) And Len(value) = 10 And Left(value, 1) = "5" Then
            targetCell.Interior.Color = xlNone
            'targetCell.CommentThreaded.Delete
        Else
            targetCell.Interior.Color = RGB(255, 0, 0)
            
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
            On Error Resume Next
            targetCell.AddCommentThreaded "Invalid value. Please ensure it is numeric, 10 characters long, and starts with '5'."
            On Error GoTo 0
        End If
    End If
        
End Sub
'Party Name
Sub ValidateAndFormatColumnE(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    
    
    rowno = rowid + 11
    
  
    Set targetCell = ActiveSheet.Cells(rowno, 5)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub
'Party Name English
Sub ValidateAndFormatColumnF(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 6)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub
'Account Name
Sub ValidateAndFormatColumnS(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 19)
    value = targetCell.value
      ' Remove existing comments
     On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub
'Account Name English
Sub ValidateAndFormatColumnT(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 20)
    value = targetCell.value
    
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub
'Site Name
Sub ValidateAndFormatColumnAF(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 32)
    value = targetCell.value
       ' Remove existing comments
     On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub
'Site Name English
Sub ValidateAndFormatColumnAG(rowid As Integer)

    Dim value As String
    Dim targetCell As Range
    Dim rowno As Integer
    rowno = rowid + 11
    Set targetCell = ActiveSheet.Cells(rowno, 33)
    value = targetCell.value
     ' Remove existing comments
    On Error Resume Next
    targetCell.CommentThreaded.Delete
    On Error GoTo 0

    If Len(value) > 80 Then
        targetCell.Interior.Color = RGB(255, 0, 0)
        ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            
        On Error Resume Next
        targetCell.AddCommentThreaded "Invalid name. Name exceed length of 80 character."
        On Error GoTo 0
    Else
       'targetCell.Interior.Color = xlNone
    End If

End Sub

'Relationship Start Date (DD-MMM-YYYY)
Sub ValidateAndFormatColumnBD(rowNumber As Integer)
    Dim rowno As Long
    rowno = rowNumber + 11
    
    Dim dateCell As Range
    Set dateCell = ActiveSheet.Cells(rowno, 56)
    On Error Resume Next
    dateCell.CommentThreaded.Delete
    On Error GoTo 0
    If (IsDate(dateCell.value)) Then
        'Debug.Print UCase(Format(dateCell.value, "DD-MMM-YYYY"))
        If UCase(Format(dateCell.value, "DD-MMM-YYYY")) = UCase(dateCell.value) Then
            dateCell.Interior.Color = xlNone
        Else
            dateCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            On Error Resume Next
            dateCell.AddCommentThreaded "Invalid Date.Please put date in 'DD-MMM-YYYY' format."
            On Error GoTo 0
          End If
    Else
        If "" <> dateCell.value Then
            dateCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            On Error Resume Next
            dateCell.AddCommentThreaded "Invalid Date.Please put date in 'DD-MMM-YYYY' format."
            On Error GoTo 0
        End If
    End If
End Sub
'Relationship End Date (DD-MMM-YYYY)
Sub ValidateAndFormatColumnBE(rowNumber As Integer)
    Dim rowno As Long
    rowno = rowNumber + 11
    
    Dim dateCell As Range
    Set dateCell = ActiveSheet.Cells(rowno, 57)
    
    On Error Resume Next
    dateCell.CommentThreaded.Delete
    On Error GoTo 0
    If (IsDate(dateCell.value)) Then
        'Debug.Print UCase(Format(dateCell.value, "DD-MMM-YYYY"))
        If UCase(Format(dateCell.value, "DD-MMM-YYYY")) = UCase(dateCell.value) Then
            dateCell.Interior.Color = xlNone
        Else
            dateCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            On Error Resume Next
            dateCell.AddCommentThreaded "Invalid Date.Please put date in 'DD-MMM-YYYY' format."
            On Error GoTo 0
          End If
    Else
        If "" <> dateCell.value Then
            dateCell.Interior.Color = RGB(255, 0, 0)
            ActiveSheet.Cells(rowno, 58).value = "Validation Failed"
            On Error Resume Next
            dateCell.AddCommentThreaded "Invalid Date.Please put date in 'DD-MMM-YYYY' format."
            On Error GoTo 0
        End If
    End If
End Sub



Private Sub LOVCheck(i As Integer, j As Integer, listRange As String)
    Dim listSheet As Worksheet
    Dim cellInList As Range
    i = i + 11
    'ActiveSheet.Cells(i, j).Interior.Color = xlNone
    On Error Resume Next
    ActiveSheet.Cells(i, j).CommentThreaded.Delete
    On Error GoTo 0

    Set listSheet = ThisWorkbook.Sheets("ListOfValues")
        For Each cellInList In listSheet.Range(listRange & listSheet.Cells(listSheet.Rows.Count, Mid(listRange, 1, 1)).End(xlUp).Row)
            If ActiveSheet.Cells(i, j).value <> "" Then
                If cellInList.value = ActiveSheet.Cells(i, j).value And ActiveSheet.Cells(i, j).value <> "" Then
                    ActiveSheet.Cells(i, j).Interior.Color = xlNone
                    On Error Resume Next
                    ActiveSheet.Cells(i, j).CommentThreaded.Delete
                    On Error GoTo 0
                    i = i - 11
                    Exit Sub
                Else
                    On Error Resume Next
                    ActiveSheet.Cells(i, j).AddCommentThreaded "Invalid Value.Please choose correct value from list of values."
                    On Error GoTo 0
                    ActiveSheet.Cells(i, j).Interior.Color = RGB(255, 0, 0)
                    ActiveSheet.Cells(i, 58).value = "Validation Failed"
                End If
            Else
                ActiveSheet.Cells(i, j).Interior.Color = xlNone
            End If
        Next cellInList
i = i - 11

End Sub

Private Sub Validate()
Worksheets("CDM Production Request Template").Activate
no_of_cols = ActiveSheet.Range("A1").CurrentRegion.Columns.Count
no_of_rows = ActiveSheet.Range("A1").CurrentRegion.Rows.Count

Set colRng = ActiveSheet.Range(ActiveSheet.Cells(12, 1), ActiveSheet.Cells(no_of_rows, no_of_cols))
Dim record As Integer
For Row = 1 To no_of_rows
'loop thru 100 rows in the sheet
    For col = 1 To no_of_cols
        temp = Row + 11
        On Error Resume Next
        ActiveSheet.Cells(temp, col).CommentThreaded.Delete "Comment deleted" 'delete threaded comment
        temp = Row - 11
    Next col 'next column
Next Row 'next row
For record = 1 To colRng.Rows.Count

' set interior color to none for all cells
Range("C" + CStr(record + 11) + ":BE" + CStr(record + 11)).Interior.Color = xlNone
'Debug.Print ("no_of_cols : " + CStr(no_of_cols))
'Debug.Print ("no_of_rows : " + CStr(no_of_rows))
    colRng.Cells(record, 58) = "Validation Successful"
    
 If (colRng.Cells(record, 2).value = "Change Relationship") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        
        'Registry Site ID
        Mandatory record, 4
        'Relation Type
        Mandatory record, 52
        'Relationship Start Date (DD-MMM-YYYY)
        Mandatory record, 56
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
        
    ElseIf (colRng.Cells(record, 2).value = "Change Address") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        
        'Registry Site ID
        Mandatory record, 4
        'Original SFDC Site ID
        Mandatory record, 35
        'Address Line 1
        Mandatory record, 36
        'City
        Mandatory record, 39
        'Postal Code
        Mandatory record, 41
        'Country
        Mandatory record, 42
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
        
    ElseIf (colRng.Cells(record, 2).value = "Change Attributes") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        'Registry Site ID
        Mandatory record, 4
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
    ElseIf (colRng.Cells(record, 2).value = "Inactivate Site") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        'Registry Site ID
        Mandatory record, 4
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
    ElseIf (colRng.Cells(record, 2).value = "Create Account") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        'Registry Party ID
        Mandatory record, 3
                'Account Name
        Mandatory record, 19
        'Account Name English
        Mandatory record, 20
        'Primary Account Flag
        Mandatory record, 21
        'Operating Unit
        Mandatory record, 22
        'Preferred Language
        Mandatory record, 23
        'Channel
        Mandatory record, 24
        'Sub Channel
        Mandatory record, 25
        'Account Class
        Mandatory record, 26
        'Customer Group
        Mandatory record, 27
        'Customer Segment 1
        Mandatory record, 28
        'Customer Segment 2
        Mandatory record, 29
        'Site Name
        Mandatory record, 32
        'Site Name English
        Mandatory record, 33
        'Site Status
        Mandatory record, 34
        'Bill To Usage
        Mandatory record, 43
        'Ship To Usage
        Mandatory record, 44
        'Store Usage
        Mandatory record, 45
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
    ElseIf (colRng.Cells(record, 2).value = "Create Party") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"

       
       ' check for Party Name
        Mandatory record, 5
        
        'check for Party Name English
        Mandatory record, 6
       
        ' check for Primary Party Flag
        Mandatory record, 7
        ' check for Party Type
        Mandatory record, 8
        ' check for Tax Registration Number
        Mandatory record, 9
        'Account Name
        Mandatory record, 19
        'Account Name English
        Mandatory record, 20
        'Primary Account Flag
        Mandatory record, 21
        'Operating Unit
        Mandatory record, 22
        'Preferred Language
        Mandatory record, 23
        'Channel
        Mandatory record, 24
        'Sub Channel
        Mandatory record, 25
        'Account Class
        Mandatory record, 26
        'Customer Group
        Mandatory record, 27
        'Customer Segment 1
        Mandatory record, 28
        'Customer Segment 2
        Mandatory record, 29
        'Site Name
        Mandatory record, 32
        'Site Name English
        Mandatory record, 33
        'Site Status
        Mandatory record, 34
        'Address Line 1
        Mandatory record, 36
        'City
        Mandatory record, 39
        'Postal Code
        Mandatory record, 41
        'Country
        Mandatory record, 42
        'Bill To Usage
        Mandatory record, 43
        'Ship To Usage
        Mandatory record, 44
        'Store Usage
        Mandatory record, 45
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)
    ElseIf (colRng.Cells(record, 2).value = "Reactivate Site") Then
        LOVCheck record, 2, "B3:B"
        LOVCheck record, 7, "D3:D"
        LOVCheck record, 8, "F3:F"
        LOVCheck record, 10, "H3:H"
        LOVCheck record, 11, "J3:J"
        LOVCheck record, 16, "L3:L"
        LOVCheck record, 17, "N3:N"
        LOVCheck record, 21, "P3:P"
        LOVCheck record, 22, "R3:R"
        LOVCheck record, 23, "T3:T"
        LOVCheck record, 24, "V3:V"
        LOVCheck record, 25, "X3:X"
        LOVCheck record, 26, "Z3:Z"
        LOVCheck record, 34, "AB3:AB"
        LOVCheck record, 42, "AD3:AD"
        LOVCheck record, 43, "AF3:AF"
        LOVCheck record, 44, "AH3:AH"
        LOVCheck record, 45, "AJ3:AJ"
        LOVCheck record, 52, "AL3:AL"
        'Registry Site ID
        Mandatory record, 4
        'Bill To Usage
        Mandatory record, 43
        'Ship To Usage
        Mandatory record, 44
        'Store Usage
        Mandatory record, 45
        ValidateAndFormatColumnD (record)
        ValidateAndFormatColumnC (record)
        ValidateAndFormatColumnE (record)
        ValidateAndFormatColumnF (record)
        ValidateAndFormatColumnS (record)
        ValidateAndFormatColumnT (record)
        ValidateAndFormatColumnAF (record)
        ValidateAndFormatColumnAG (record)
        ValidateAndFormatColumnAI (record)
        ValidateAndFormatColumnAT (record)
        ValidateAndFormatColumnAU (record)
        ValidateAndFormatColumnAW (record)
        ValidateAndFormatColumnAX (record)
        ValidateAndFormatColumnBA (record)
        ValidateAndFormatColumnBB (record)
        ValidateAndFormatColumnBD (record)
        ValidateAndFormatColumnBE (record)

    End If
    Next record
	MsgBox ("Validation complete.")
End Sub