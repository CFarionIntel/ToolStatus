Attribute VB_Name = "Main_WOPR_Update"
Sub Main_WOPR()
Attribute Main_WOPR.VB_ProcData.VB_Invoke_Func = "w\n14"
    Const DEV As Boolean = False
    
    Application.ScreenUpdating = False
    'Collect the Sheet Names and columns which comprise the location of the WOPRs
    Dim WorkBook_Component_Names As Collection
    Set WorkBook_Component_Names = Tool_Status_WorkBook_Setup
    
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
        
        Record_WOPR_Results Change_Report
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Function Process_for_Each_Entry(ByRef WorkBook_Info As Collection, ByRef Entries As Collection) As Collection
    Dim Already_Checked_Collection As New Collection
    Dim Entry_PreChecks_Passing As Boolean
    Dim Change_Report_Line As String
    Dim Change_Report As New Collection
    Dim Change_Report_Open As New Collection
    Dim Change_Report_Closed As New Collection
    
    With Change_Report
        .Add Change_Report_Open
        .Add Change_Report_Closed
    End With
    
    For Entry = 1 To Entries.Count
        ' Check that the current entry is free of errors
        ' This function will also set the cell value of the current entry according to the Tool Status Page in the WorkBook_Info Collection
        Entry_PreChecks_Passing = Entry_PreChecks(WorkBook_Info, Entries(Entry))
        If Not Entry_PreChecks_Passing Then
            GoTo NextEntry
        End If
        
        'The WOPR link will be created or removed from the Tool Status page
        'Output will report back the change to the sheet
        Change_Report_Line = Add_or_Remove_WOPR_Links(WorkBook_Info, Entries(Entry))
        
        If Change_Report_Line <> "" Then
            If Entries(Entry).Sts = "Closed" Then
                Change_Report_Closed.Add Change_Report_Line
            Else
                Change_Report_Open.Add Change_Report_Line
            End If
        End If
NextEntry:
        Progress = PercentDone(CLng(Entry), Entries.Count, "Boca Juniors sos mi vida!")
        Application.StatusBar = "Processing WOPR: " & Entry & " / " & Entries.Count & " " & Progress
        'Application.StatusBar = "Processing WOPR: " & Entry & " / " & Entries.Count
    Next Entry
    
    Set Process_for_Each_Entry = Change_Report
End Function

Private Function Add_or_Remove_WOPR_Links(ByRef WorkBook_Info As Collection, ByRef WOPR_Entry As WOPR) As String
    Dim ret As String
    ret = ""
    
    If WOPR_Entry.Sts = "Closed" Then
        ret = Multiple_WOPRs_Close_Routine(WorkBook_Info, WOPR_Entry)
    Else 'Not "Closed"
        ret = Multiple_WOPRs_Open_Routine(WorkBook_Info, WOPR_Entry)
    End If
        
    Add_or_Remove_WOPR_Links = ret
End Function

Private Function Multiple_WOPRs_Open_Routine(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR) As String
    Dim ret As String
    ret = ""
    
    Const WOPR_Link_Address As String = "https://rf3-apps-fuzion.rf3prod.mfg.intel.com/EditWorkOrderPage.aspx?WorkOrderId="
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_First_WOPR_Column As Long
    DashBoard_First_WOPR_Column = WorkBook_Info(7)
    
    Dim ENTITY_CELL As String
    ENTITY_CELL = WorkBook_Info(WorkBook_Info.Count)
    
    Dim Number_of_WOPRs_in_Row As Long
    Number_of_WOPRs_in_Row = Determine_Number_of_WOPRs_in_Row(WorkBook_Info)
    
    For col = DashBoard_First_WOPR_Column To (DashBoard_First_WOPR_Column + Number_of_WOPRs_in_Row - 1)
        current_val = Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(ENTITY_CELL).Row, col)
        If current_val = WOPR_Entry.ID Then
            Exit Function
        End If
    Next col
    Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(ENTITY_CELL).Row, DashBoard_First_WOPR_Column + Number_of_WOPRs_in_Row) = WOPR_Entry.ID
    Cell2Link = Global_Functions.Column2Letter(DashBoard_First_WOPR_Column + Number_of_WOPRs_in_Row) & Range(ENTITY_CELL).Row
    Link_text = WOPR_Entry.ID
    Worksheets(ToolSts_DashBoard_SheetName).Hyperlinks.Add Anchor:=Range(Cell2Link), Address:=WOPR_Link_Address & WOPR_Entry.ID, TextToDisplay:=CStr(WOPR_Entry.ID)
    ret = "WO# " & WOPR_Entry.ID & " for " & WOPR_Entry.Entity & " is " & WOPR_Entry.Sts
    Multiple_WOPRs_Open_Routine = ret
End Function

Private Function Multiple_WOPRs_Close_Routine(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR) As String
    Dim ret As String
    ret = ""
    
    Dim Last_WOPR_Column As Long
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_First_WOPR_Column As Long
    DashBoard_First_WOPR_Column = WorkBook_Info(7)
    
    Dim ENTITY_CELL As String
    ENTITY_CELL = WorkBook_Info(WorkBook_Info.Count)
    
    Dim Number_of_WOPRs_in_Row As Long
    Number_of_WOPRs_in_Row = Determine_Number_of_WOPRs_in_Row(WorkBook_Info)
    
    If Number_of_WOPRs_in_Row = 0 Then
        Exit Function
    End If
    
    Last_WOPR_Column = DashBoard_First_WOPR_Column + Number_of_WOPRs_in_Row - 1
    
    For col = DashBoard_First_WOPR_Column To Last_WOPR_Column
        current_val = Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(ENTITY_CELL).Row, col)
        If current_val = WOPR_Entry.ID Then
            Dim First_Column_of_Cut As Long
            First_Column_of_Cut = col + 1
            
            Dim Last_Column_of_Cut As Long
            Last_Column_of_Cut = Last_WOPR_Column
            
            'Case of 1 WOPR in the row
            If Last_Column_of_Cut < First_Column_of_Cut Then
                Last_Column_of_Cut = First_Column_of_Cut
            End If
            
            Range_of_Cut = Global_Functions.Column2Letter(First_Column_of_Cut) & Range(ENTITY_CELL).Row & ":" & Global_Functions.Column2Letter(Last_Column_of_Cut) & Range(ENTITY_CELL).Row
            ActiveWorkbook.Sheets(ToolSts_DashBoard_SheetName).Range(Range_of_Cut).Cut Destination:=Sheets(ToolSts_DashBoard_SheetName).Range(Cells(Range(ENTITY_CELL).Row, col).Address)
            ret = "WO# " & WOPR_Entry.ID & " for " & WOPR_Entry.Entity & " is " & WOPR_Entry.Sts
            Multiple_WOPRs_Close_Routine = ret
            Exit Function
        End If
    Next col
    Multiple_WOPRs_Close_Routine = ret
End Function

Private Function Determine_Number_of_WOPRs_in_Row(ByRef WorkBook_Info As Collection) As Long
    Dim ret As Long
    ret = 0
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_First_WOPR_Column As Long
    DashBoard_First_WOPR_Column = WorkBook_Info(7)
    
    Dim ENTITY_CELL As String
    ENTITY_CELL = WorkBook_Info(WorkBook_Info.Count)
    
    Dim Value_at_First_WOPR_Column As String
    Value_at_First_WOPR_Column = Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(ENTITY_CELL).Row, DashBoard_First_WOPR_Column).Value
    
    Dim Last_WOPR_Column As Long
    Last_WOPR_Column = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(ENTITY_CELL).Row, DashBoard_First_WOPR_Column).End(xlToRight).Column
    
    If Value_at_First_WOPR_Column <> "" Then
        If Last_WOPR_Column > 16383 Then
            ret = 1
        Else
            ret = Last_WOPR_Column - DashBoard_First_WOPR_Column + 1
        End If
    End If
    Determine_Number_of_WOPRs_in_Row = ret
End Function

Private Sub Add_to_Change_Report(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR, ByRef Change_Report As Collection)
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim Message As String
    Message = ""
    
End Sub

Private Function Find_Entity_Cell(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR) As String
    Dim ret As String
    ret = "SKIP"
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = WorkBook_Info(3)
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = WorkBook_Info(2)

    Dim first_row As Long
    first_row = 1
    
    Dim Current_Row As Long
    Dim current_value As String
    
    While first_row < DashBoard_Last_Row And (first_row + 1) <> DashBoard_Last_Row
        Current_Row = (first_row + DashBoard_Last_Row) \ 2
        current_value = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Cells(Current_Row, DashBoard_Entity_Column)
        If Simplify(current_value) = Simplify(WOPR_Entry.Entity) Then
            first_row = DashBoard_Last_Row + 1
        ElseIf Simplify(current_value) < Simplify(WOPR_Entry.Entity) Then
            first_row = Current_Row
        ElseIf Simplify(current_value) > Simplify(WOPR_Entry.Entity) Then
            DashBoard_Last_Row = Current_Row
        End If
    Wend
    
    If (Simplify(current_value) = Simplify(WOPR_Entry.Entity)) Then
        ret = Global_Functions.Column2Letter(DashBoard_Entity_Column) & Current_Row
    ElseIf (Cells(Current_Row + 1, DashBoard_Entity_Column) = WOPR_Entry.Entity) Then
        ret = Global_Functions.Column2Letter(DashBoard_Entity_Column) & Current_Row + 1
    Else
        'Not in Tool Status Page
    End If
    WorkBook_Info.Remove WorkBook_Info.Count
    WorkBook_Info.Add ret
    Find_Entity_Cell = ret
End Function

Public Function Simplify(Entity As String) As String
    Dim ret As String
    
    comp = Split(Entity, "_")
    
    If UBound(comp) <> 1 Then
        Simplify = Entity
        Exit Function
    End If
    
    ret = comp(0) & "_" & Right(comp(1), 1)
    
    Simplify = ret
End Function

Private Function Entry_PreChecks(ByRef WorkBook_Info As Collection, ByVal WOPR_Entry As WOPR) As Boolean
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
    
    Entity_Cell_Location = Find_Entity_Cell(WorkBook_Info, WOPR_Entry)
    If Entity_Cell_Location = "SKIP" Then
        Debug.Print ("WOPR not in Output: " & WOPR_Entry.Entity)
        Entry_PreChecks = ret
        Exit Function
    End If
    
    ret = True
    Entry_PreChecks = ret
End Function

Function Create_Latest_WOPR_Collection(Input_Option As Long)
    Dim Entries As New Collection
    
    If Input_Option = 1 Then
        'Create Collection from Sheet Values
        Application.StatusBar = "Using Manual Input"
    ElseIf Input_Option = 2 Then
        Application.StatusBar = "Reading file on local drive"
        Dim STRING_SQL_OUTPUT_FILE As String
        STRING_SQL_OUTPUT_FILE = "C:\Users\cfarion\AppData\Local\Temp\SQLPathFinder_Temp\out_14118.tab"
        
        Dim Text As String
        Text = Global_Functions.Read_SQL_Output_File(STRING_SQL_OUTPUT_FILE)
        
        Populate_Entry_Collection Entries, Text
    ElseIf Input_Option = 3 Then
        Application.StatusBar = "Running SQL through Excel"
        SQL_Script_Through_Excel Entries
    Else
        'Nothing Selected or Error
        Application.StatusBar = "Not running anything"
    End If

    Set Create_Latest_WOPR_Collection = Entries
End Function

Private Sub SQL_Script_Through_Excel(ByRef coll As Collection)
    On Error GoTo ErrHandler

    Dim DataSource As String
    DataSource = "D1D_PROD_XEUS"

    Dim sQuery As String
    sQuery = ActiveWorkbook.Sheets("SQL_INPUT").Cells(3, 2).Value

    ' Initialize UniqeClientHelper class
    Dim helper As Object
    Set helper = CreateObject("Intel.FabAuto.ESFW.DS.UBER.UniqeClientHelper")
    helper.ConnectionString = "Site=BEST;Metadata=Excel sample;DataSource=" & DataSource

    ' Run the query and get the handle to the IUberTable instance which contains the query result
    Dim uberTable As Object
    Set uberTable = helper.GetUberTable(sQuery)

    ' Convert IUberTable to ADODB.Recordset (ideal for VBScript usage)
    Dim recordSet As Object
    Set recordSet = uberTable.ConvertToRecordset()

    ' Check if the recordset is empty
    If recordSet.EOF And recordSet.BOF Then
        MsgBox "No records found."
        Exit Sub
    End If

    ' Initialize the collection
    Dim Entry As WOPR
    Dim field As Object
    
    ' Loop through the recordset and add each row to the collection
    Do While Not recordSet.EOF
        Set Entry = New WOPR
        For Each field In recordSet.Fields
            fieldName = field.Name
            If fieldName = "TOOL_NAME" Then
                Entry.Entity = field.Value
            ElseIf fieldName = "AVAILABILITY" Then
                'entry.Availability = field.Value
            ElseIf fieldName = "STATE" Then
                Entry.State = field.Value
            ElseIf fieldName = "CEID" Then
                Entry.CEID = field.Value
            ElseIf fieldName = "WORKORDER_ID" Then
                Entry.ID = field.Value
            ElseIf fieldName = "STATUS" Then
                Entry.Sts = field.Value
            ElseIf fieldName = "PRIORITY_ID" Then
                Entry.Prio = field.Value
            ElseIf fieldName = "LAST_UPDATED_DATE" Then
                Entry.LastUpdated = field.Value
            ElseIf fieldName = "WORKORDER_DESC" Then
                Entry.Description = field.Value
            Else
                'Error / Do Nothing
            End If
        Next field
        coll.Add Entry
        recordSet.MoveNext
    Loop

    ' Clean up
    recordSet.Close
    Set recordSet = Nothing
    Set helper = Nothing
    Set uberTable = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub Populate_Entry_Collection(ByRef coll As Collection, ByVal File_Text As String)
    Dim Entry As WOPR
    
    Dim Name_index As Long
    Dim Availability_index As Long
    Dim State_index As Long
    Dim CEID_index As Long
    Dim Prio_index As Long
    Dim WOPR_ID_index As Long
    Dim WOPR_Sts_index As Long
    Dim Last_Updated_index As Long
    Dim WOPR_Description_index As Long
    
    Instances = Split(File_Text, Chr(10))
    For Instance = 0 To UBound(Instances) - 1
        Set Entry = New WOPR
        Components = (Split(Instances(Instance), Chr(9)))
        If Instance = 0 Then 'Determine which column is which variable. Good for if someone changes the output columns of SQL script
            For i = 0 To UBound(Components)
                If Components(i) = "TOOL_NAME" Then
                    Name_index = i
                ElseIf Components(i) = "AVAILABILITY" Then
                    Availability_index = i
                ElseIf Components(i) = "STATE" Then
                    State_index = i
                ElseIf Components(i) = "CEID" Then
                    CEID_index = i
                ElseIf Components(i) = "PRIORITY_ID" Then
                    Prio_index = i
                ElseIf Components(i) = "WORKORDER_ID" Then
                    WOPR_ID_index = i
                ElseIf Components(i) = "STATUS" Then
                    WOPR_Sts_index = i
                ElseIf Components(i) = "LAST_UPDATED_DATE" Then
                    Last_Updated_index = i
                ElseIf Components(i) = "WORKORDER_DESC" Then
                    WOPR_Description_index = i
                Else
                    'Error / Do Nothing
                End If
            Next i
        Else
            If UBound(Components) > 0 Then
                With Entry
                    On Error Resume Next
                        .ID = CLng(Components(WOPR_ID_index))
                    .Sts = Components(WOPR_Sts_index)
                    .Entity = Components(Name_index)
                    .CEID = Components(CEID_index)
                    On Error Resume Next
                        .Prio = Components(Prio_index)
                    '.CreatedDate = ""
                    .LastUpdated = Components(Last_Updated_index)
                    .State = Components(State_index)
                End With
                     
                coll.Add Entry
            End If
        End If
    Next Instance
End Sub


Private Function WorkBook_Setup_WOPR() As Collection
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const ToolSts_DashBoard_SheetName As String = "Tool Status"
    
    'Prep Output Sheet
    On Error Resume Next
        ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).ShowAllData
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = Global_Functions.Find_Col("Entity", ToolSts_DashBoard_SheetName)
    
    Dim DashBoard_First_WOPR_Column As Long 'Today's comments column
    DashBoard_First_WOPR_Column = Global_Functions.Find_Col("WOPR ID", ToolSts_DashBoard_SheetName)
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(Global_Functions.Column2Letter(DashBoard_Entity_Column) & "65535").End(xlUp).Row

    Entity_Column_Range = Global_Functions.Column2Letter(DashBoard_Entity_Column) & "1:" & Global_Functions.Column2Letter(DashBoard_Entity_Column) & DashBoard_Last_Row
        
    'This MUST be in alphabetical order or else the algorithms do not work
    ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort.SortFields.Add2 Key:=Range(Entity_Column_Range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    With ret_Component_Names
        .Add ToolSts_DashBoard_SheetName
        .Add DashBoard_Entity_Column
        .Add DashBoard_First_WOPR_Column
        .Add DashBoard_Last_Row
        .Add TEMP_CELL_HOLDER
    End With

    
    Set WorkBook_Setup_WOPR = ret_Component_Names
End Function

Sub WOPR_Links()
Attribute WOPR_Links.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim Link_text As String
    
    If Max_on_Same_Entity = 0 Then
        Max_on_Same_Entity = 2
    End If
    
    Const STRING_OUTPUT_SHEET As String = "Tool Status"
    
    OUTPUT_WOPR_COL = Global_Functions.Find_Col("WOPR ID", STRING_OUTPUT_SHEET)
    
    
    
    For col = OUTPUT_WOPR_COL To OUTPUT_WOPR_COL + Max_on_Same_Entity
        For Row = 2 To 700
            WOPR_ID = Cells(Row, col)
            If WOPR_ID <> "" Then
                Link_text = WOPR_ID
                STRING_chamber_col_letter = Split(Cells(Row, col).Address, "$")(1)
                Cell2Link = STRING_chamber_col_letter & Row
                ActiveSheet.Hyperlinks.Add Anchor:=Range(Cell2Link), Address:="https://rf3-apps-fuzion.rf3prod.mfg.intel.com/EditWorkOrderPage.aspx?WorkOrderId=" & WOPR_ID, TextToDisplay:=Link_text
            End If
        Next Row
    Next col
End Sub

Sub Record_WOPR_Results(ByRef Report As Collection)
    Num_of_Open = Report(1).Count
    Num_of_Closed = Report(2).Count
    
    Open_String = ""
    Closed_String = ""
    
    If Num_of_Open = 0 Then
        Open_String = "No New Non-Closed Work Orders"
    Else
        For i = 1 To Num_of_Open - 1
            Open_String = Open_String & Report(1)(i) & Chr(10)
        Next i
        Open_String = Open_String & Report(1)(i)
    End If
    
    If Num_of_Closed = 0 Then
        Closed_String = "No Closed Work Orders Since Last Query"
    Else
        For i = 1 To Num_of_Closed - 1
            Closed_String = Closed_String & Report(2)(i) & Chr(10)
        Next i
        Closed_String = Closed_String & Report(2)(i)
    End If
    
    Debug.Print ("End > " & Format(Time, "hh:mm.ss"))
    Time_Format = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
    rowh = ActiveWorkbook.Sheets("Change Report").Range("A1812").End(xlUp).Row + 1
    With Sheets("Change Report")
        .Cells(rowh, 1) = Time_Format
        .Cells(rowh, 2) = Open_String
        .Cells(rowh, 3) = Closed_String
    End With
    If (Num_of_Open + Num_of_Closed) > 18 Then
        Open_String = "Multiple Changes above Message Box Limit. Please check the Change Report Tab."
        Closed_String = ""
    Else
        Open_String = "Change Report: " & Chr(10) & Chr(10) & Open_String
    End If
    out = MsgBox(Open_String & Chr(10) & Closed_String, 0, Time_Format, 0, 0)
End Sub

Sub FDC_E3_Message_Output()
    Dim WorkBook_Input_Parameters As Collection
    Set WorkBook_Input_Parameters = Abort_Input_WorkBook_Setup
    
    Dim ret As String
    
    Dim Ent_WOPR As New WOPR
    Dim Module As String
    
    Dim Abort_Setup_Sheetname As String
    Abort_Setup_Sheetname = WorkBook_Input_Parameters(1)
    
    Const Title_Format As String = "[..::..] POR Lot Abort Recovery - HB? Lot ..::.. - ..::.."
    
    Dim E3_Message As String
    E3_Message = ActiveWorkbook.Sheets(Abort_Setup_Sheetname).Cells(WorkBook_Input_Parameters(8), WorkBook_Input_Parameters(2)).Value
    
    Dim Error_for_Title As String
    Error_for_Title = ""
    
    Parts = Split(Title_Format, "..::..")
    
    s0 = Split(E3_Message, ",")
    Lot = Split(s0(0), ":")(1)
    If Mid(Lot, 5, 1) = "T" Then
        Lot = "DCS"
    End If
    Ent_WOPR.Entity = Split(s0(1), ":")(1)
    Error_comp = Split(s0(3), " ")
    ErrorMsg = Split(Error_comp(1), ":")
    part0 = Split(ErrorMsg(1), "/")(0)
    part1 = Split(ErrorMsg(1), "/")(1)
    part2 = Split(ErrorMsg(1), "/")(2)
    
    If part1 = "CustomEquation" Then
        Final_Error = part2
    Else
        Final_Error = part1 & "/" & part2
    End If
    
    Dim WorkBook_Dashboard_Parameters As Collection
    Set WorkBook_Dashboard_Parameters = Tool_Status_WorkBook_Setup
    
    ret = Find_Entity_Cell(WorkBook_Dashboard_Parameters, Ent_WOPR)
    
    If ret <> "SKIP" Then
        Module = Sheets(WorkBook_Dashboard_Parameters(1)).Cells(Range(ret).Row, 1)
        Bay = Sheets(WorkBook_Dashboard_Parameters(1)).Cells(Range(ret).Row, 6)
    End If
    
    AMF4_Route = ""
    AMF4_Entity = Ent_WOPR.Entity
    AMF4_Event = "4P" & Right(AMF4_Entity, 1) & "_ETCH_TEST"
    AMF4_Chamber = Right(Ent_WOPR.Entity, 3)
    
    Select Case Module
        Case "DE-NIT-UL", "DE-MT-UL", "DE-VIA-UL", "GTOcz", "TVOcz"
            AMF4_Route = "ECTW.310"
            
        Case "DE-HM2-SLM"
            AMF4_Route = "ECTW.407"
        
        Case "DE-HM2-GRT"
            AMF4_Route = "ECTW.406"
        
        Case "DE-GTOcw"
            AMF4_Route = "ECTW.U64"
            
        Case "DE-TVO-TL", "DE-TVO-TLCP7"
            AMF4_Route = "ENTW.Y56"
            
        Case "DE-GTO-CS"
            AMF4_Route = "ECTW.V72"
            AMF4_Event = "AMF4_TACOCAT"
        
    End Select
    
    'AMF4_Entity = Ent_WOPR.Entity
    'AMF4_Event = "4P" & Right(AMF4_Entity, 1) & "_ETCH_TEST"
    AMF4_Recipe = "PX/P*"
    OutputMsg = Parts(0) & Module & Parts(1) & Lot & Parts(2) & Final_Error & Chr(10) & Chr(10) & E3_Message
    OutputMsg1 = Parts(0) & Module & "] " & Final_Error & " Recovery" & Chr(10) & Chr(10) & E3_Message
    AMF4 = "Route: " & AMF4_Route & Chr(10) & "Entity: " & AMF4_Entity & Chr(10) & "Event: " & AMF4_Event & Chr(10) & "Recipe: " & AMF4_Recipe & Chr(10) & "Chamber: " & AMF4_Chamber & Chr(10) & "Lot: " & Chr(10) & "Slots: Any 3"
    Worksheets(Abort_Setup_Sheetname).Cells(47, 1) = OutputMsg
    Worksheets(Abort_Setup_Sheetname).Cells(48, 1) = "Hi team, here is WO# 3811495 for the non-HB and non-CQT abort on " & AMF4_Entity & " (" & Bay & "). Thank you!"
    Worksheets(Abort_Setup_Sheetname).Cells(49, 1) = OutputMsg1
    Worksheets(Abort_Setup_Sheetname).Cells(50, 1) = AMF4
    
    
End Sub

