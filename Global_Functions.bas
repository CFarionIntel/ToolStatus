Attribute VB_Name = "Global_Functions"
Public Function Read_SQL_Output_File(file_path_and_name As String) As String
    Dim Text As String
    
    Text = ""
    Open file_path_and_name For Input As #1
        Do Until EOF(1)
            Line Input #1, textline
            Text = Text & textline & Chr(10)
        Loop
        Close #1
        Read_SQL_Output_File = Text
End Function

Public Function Column2Letter(col As Long) As String
    Dim ret As String
    ret = Split(Cells(1, col).Address, "$")(1)
    Column2Letter = ret
End Function

Public Function Find_Col(Header As String, Optional SHEET_NAME As String, Optional STARTING_ROW As Long, Optional STARTING_COL As Long) As Long
    Dim sheet2use As String
    Dim row2start As Long
    row2start = 1
    Dim col2start As Long
    col2start = 1
    
    If SHEET_NAME = "" Then
        sheet2use = ActiveSheet.Name
    ElseIf SHEET_NAME <> "" Then
        sheet2use = ActiveWorkbook.Worksheets(SHEET_NAME).Name
    End If
    
    If STARTING_ROW <> 0 Then
        row2start = STARTING_ROW
    End If
    
    If STARTING_COL <> 0 Then
        col2start = STARTING_COL
    End If
    
    last_col = CLng(Split(STRING_get_last_row_and_col(sheet2use, row2start, col2start), ";")(1))
    
    For current_col = 1 To last_col
        current_val = ActiveWorkbook.Worksheets(sheet2use).Cells(row2start, (col2start - 1) + current_col)
        If LCase(current_val) = LCase(Header) Then
            Exit For
        End If
    Next current_col
    
    If current_col > last_col Then
        current_col = 0
    End If
    Find_Col = current_col
End Function

Public Sub Create_WOPR_Links(Optional ByVal Max_on_Same_Entity As Long)
    Dim Link_text As String
    
    If Max_on_Same_Entity = 0 Then
        Max_on_Same_Entity = 2
    End If
    
    Const STRING_OUTPUT_SHEET As String = "Tool Status"
    Const WOPR_Hyperlink As String = "https://rf3-apps-fuzion.rf3prod.mfg.intel.com/EditWorkOrderPage.aspx?WorkOrderId="
    
    OUTPUT_WOPR_COL = Find_Col("WOPR ID", STRING_OUTPUT_SHEET)
    
    For col = OUTPUT_WOPR_COL To OUTPUT_WOPR_COL + Max_on_Same_Entity
        For Row = 2 To 700
            WOPR_ID = Cells(Row, col)
            If WOPR_ID <> "" Then
                Link_text = WOPR_ID
                STRING_chamber_col_letter = Split(Cells(Row, col).Address, "$")(1)
                Cell2Link = STRING_chamber_col_letter & Row
                ActiveSheet.Hyperlinks.Add Anchor:=Range(Cell2Link), Address:=WOPR_Hyperlink & WOPR_ID, TextToDisplay:=Link_text
            End If
        Next Row
    Next col
    MsgBox ("Links completed")
End Sub

Public Sub a2l(String2Add As String)
    'a2l -> Add to Log
    'Assume Logs are a Public and Global Collection
    Dim retString As String
    Const delimiter As String = ";"
    Const txtSplit As String = ">"
    
    LogDate = Format(Date, "yyyy/mm/dd")
    LogTime = Format(Time, "hh:mm.ss")
    
    retString = LogDate & delimiter & LogTime & " " & txtSplit & " " & String2Add
    
    
    LogCollection.Add retString
End Sub

