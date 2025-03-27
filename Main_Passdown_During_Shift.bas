Attribute VB_Name = "Main_Passdown_During_Shift"
Sub Display_WOPRs_Updated_During_Shift()
    Const DEV As Boolean = True
    
    Debug.Print ("Start WOPRs This Shift > " & Format(Time, "hh:mm.ss"))
    Application.ScreenUpdating = False
    Application.StatusBar = True
    
    'Collect the Sheet Names and columns which comprise the location of the WOPRs
    Dim WorkBook_Component_Names As Collection
    Set WorkBook_Component_Names = Updated_WOPRs_During_Shift_WorkBook_Setup
    
    'Select how you want to collect the new WOPR Update
    Dim WOPR_Input_Selection As Long
    WOPR_Input_Selection = ActiveWorkbook.Sheets("Settings").Cells(1, 12)
    
    'Create a collection of the latest Tool Status from the Input
    Dim Entries As Collection
    Set Entries = Create_Latest_WOPR_Collection(WOPR_Input_Selection)
    
    If Entries.Count <> 0 Then
        'Processing for each entry in Query. This will output a Change Report
        Dim Change_Report As New Collection
        Set Change_Report = Process_for_Each_Entry(WorkBook_Component_Names, Entries)
        
        'Record_WOPR_Results Change_Report
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print ("End WOPRs This Shift > " & Format(Time, "hh:mm.ss"))
End Sub

Private Function Process_for_Each_Entry(ByRef WorkBook_Info As Collection, ByRef Entries As Collection) As Collection
    Dim Already_Checked_Collection As New Collection
    Dim Entry_PreChecks_Passing As Boolean
    Dim Change_Report_Line As String
    Dim Change_Report As New Collection
    Dim Current_Row As Long
    Current_Row = WorkBook_Info(2)
    Const Days_Back As Long = 1
    Dim currentDate As String
    currentDate = Format(Date - Days_Back, "yyyy-mm-dd")
    Const Cutoff_Time As String = "19:30:00"
    Dim Beginning_of_Shift_Time As String
    Beginning_of_Shift_Time = currentDate & " " & Cutoff_Time
    
    For Entry = 1 To Entries.Count
        ' Check that the current entry is free of errors
        ' This function will also set the cell value of the current entry according to the Tool Status Page in the WorkBook_Info Collection
        Entry_PreChecks_Passing = Entry_PreChecks(WorkBook_Info, Entries(Entry), Beginning_of_Shift_Time)
        If Not Entry_PreChecks_Passing Then
            GoTo NextEntry
        End If
        
        With ActiveWorkbook.Sheets(WorkBook_Info(1))
            .Cells(Current_Row, WorkBook_Info(3)) = Entries(Entry).Entity
            .Cells(Current_Row, WorkBook_Info(4)) = Entries(Entry).CEID
            .Cells(Current_Row, WorkBook_Info(5)) = Entries(Entry).State
            .Cells(Current_Row, WorkBook_Info(6)) = Entries(Entry).ID
            .Cells(Current_Row, WorkBook_Info(7)) = Entries(Entry).Sts
            .Cells(Current_Row, WorkBook_Info(8)) = Entries(Entry).Prio
            .Cells(Current_Row, WorkBook_Info(9)) = Entries(Entry).LastUpdated
            .Cells(Current_Row, WorkBook_Info(10)) = Entries(Entry).Description
        End With
        Current_Row = Current_Row + 1
NextEntry:
        Application.StatusBar = "Processing chamber: " & Entry & " / " & Entries.Count
    Next Entry
    
    Set Process_for_Each_Entry = Change_Report
End Function

Private Function Entry_PreChecks(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR, Beginning_of_Shift_Time As String) As Boolean
    Dim ret As Boolean
    ret = False
        
    Dim Legitimate_WOPR As Boolean
    Legitimate_WOPR = WOPR_Entry.Verify_WOPR
    If Legitimate_WOPR <> True Then
        Debug.Print ("Not a real entity or template for WO# " & WOPR_Entry.ID)
        Entry_PreChecks = ret
        Exit Function
    End If
    
    Components = Split(WOPR_Entry.Entity, "_")
    temp = ""
    If Components(0) = "SF" Then
        For i = 1 To UBound(Components)
            temp = temp & Components(i) & "_"
        Next i
        WOPR_Entry.Entity = Left(temp, Len(temp) - 1) 'Erase the final '_' character
    End If
    
    'Look at the Date and determine if that can be added to the list
    Last_Update = WOPR_Entry.LastUpdated
    
    If Last_Update < Beginning_of_Shift_Time Then
        Debug.Print ("WO# " & WOPR_Entry.ID & " for " & WOPR_Entry.Entity & " was not updated this shift.")
    Else
        Debug.Print ("Add WO# " & WOPR_Entry.ID & " for " & WOPR_Entry.Entity & " to the list!")
        ret = True
    End If
    
    Entry_PreChecks = ret
End Function

