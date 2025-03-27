Attribute VB_Name = "Main_ToolStatus"
Public LogCollection As Collection
Public CommentCollection As Collection

Sub Main_ToolSts()
Attribute Main_ToolSts.VB_Description = "Run the Main Tool Status Macro"
Attribute Main_ToolSts.VB_ProcData.VB_Invoke_Func = "S\n14"
    Const DEV As Boolean = False
    
    InitializeLogCollection
    InitializeCommentCollection
    
    Debug.Print ("Start Main > " & Format(Time, "hh:mm.ss"))
    a2l "Start Tool Status Main"
    
    Application.ScreenUpdating = False
    'Collect the Sheet Names and columns which comprise the Dashboard
    Dim WorkBook_Component_Names As New Collection
    
    Dim Dashboard_WorkBook_Parameters As Collection
    Set Dashboard_WorkBook_Parameters = Tool_Status_WorkBook_Setup
    
    Dim ToolStsHistory_WorkBook_Parameters As Collection
    Set ToolStsHistory_WorkBook_Parameters = Tool_Status_History_WorkBook_Setup
    
    Dim Change_Report_WorkBook_Parameters As Collection
    Set Change_Report_WorkBook_Parameters = Change_Report_WorkBook_Setup
    
    With WorkBook_Component_Names
        .Add Dashboard_WorkBook_Parameters
        .Add ToolStsHistory_WorkBook_Parameters
        .Add Change_Report_WorkBook_Parameters
    End With
    
    'Select how you want to collect the new Tool Status Update
    Dim ToolSts_Input_Selection As Long
    ToolSts_Input_Selection = ActiveWorkbook.Sheets("Settings").Cells(1, 1)
    
    'Create a collection of the latest Tool Status from the Input
    Dim Entries As Collection
    Set Entries = Create_Latest_ToolSts_Collection(ToolSts_Input_Selection)
    
    If Entries.Count <> 0 Then
        Move_Comments_for_New_Day WorkBook_Component_Names
        
        'Processing for each entry in Query. This will output a Change Report
        Dim Change_Report As New Collection
        Set Change_Report = Process_for_Each_Entry(Dashboard_WorkBook_Parameters, Entries)
        
        'Take Change Report and output to Change Report Sheets
        Record_Changes WorkBook_Component_Names, Change_Report
    End If
    
    Debug.Print ("End Main > " & Format(Time, "hh:mm.ss"))
    
    a2l "End Tool Status Main"
    SaveLogsToTextFile
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub Move_Comments_for_New_Day(ByRef WorkBook_Info As Collection)
    
    If WorkBook_Info(3)(2) <> 1 Then
        Last_Update = ActiveWorkbook.Worksheets(WorkBook_Info(3)(1)).Cells(WorkBook_Info(3)(2) - 1, WorkBook_Info(3)(3))
    Else
        Last_Update = "12/18/2022 - 07:00:00"
    End If
    Latest_Update_Date = Split(Last_Update, " ")(0)
    Latest_Update_Time = Split(Last_Update, " ")(2)
    
    Today = Date
    Time_Now = Time

    deltaDate = Today - DateValue(Latest_Update_Date)
    deltaTime = Time_Now - TimeValue(Latest_Update_Time)

    Minutes_Difference = 1440 * (deltaDate + deltaTime)
    Hours_for_New_Day = 60 * 6
    
    If Minutes_Difference > Hours_for_New_Day Then
        Collect_and_Save_Comments_for_the_Day WorkBook_Info
        SaveCommentsToTextFile
        RANGE_output2paste = Global_Functions.Column2Letter(CLng(WorkBook_Info(1)(5))) & "2:" & Global_Functions.Column2Letter(CLng(WorkBook_Info(1)(5))) & WorkBook_Info(1)(3)
        ActiveWorkbook.Sheets(WorkBook_Info(1)(1)).Range(RANGE_output2paste).Cut Destination:=Sheets(WorkBook_Info(1)(1)).Range(Global_Functions.Column2Letter(WorkBook_Info(1)(5) + 1) & "2")
        a2l "New Day! Comments moved over."
    End If
End Sub

Sub Collect_and_Save_Comments_for_the_Day(ByRef WorkBook_Info As Collection)
    ' 1 ToolSts_DashBoard_SheetName
    ' 2 DashBoard_Entity_Column
    ' 5 DashBoard_Comments_Column
    
    Dim DashBoard_Comments_Column As Long
    DashBoard_Comments_Column = WorkBook_Info(1)(5)
    DashBoard_Entity_Column = WorkBook_Info(1)(2)
    
    Dim Cell_Value As String
    Dim return_line As String
    return_line = ""
    
    For i_row = 2 To WorkBook_Info(1)(3)
        Cell_Value = Sheets(WorkBook_Info(1)(1)).Cells(i_row, DashBoard_Comments_Column).Value
        If Cell_Value <> "" Then
            Entity_val = Sheets(WorkBook_Info(1)(1)).Cells(i_row, DashBoard_Entity_Column).Value
            return_line = Entity_val & " - " & Cell_Value
            CommentCollection.Add return_line
        End If
    Next i_row
End Sub

Private Function Process_for_Each_Entry(ByRef WorkBook_Info As Collection, ByRef Entries As Collection) As Collection
    Dim Already_Checked_Collection As New Collection
    Dim Entry_PreChecks_Passing As Boolean
    Dim Change_Report As New Collection
    Dim Progress As String
    Progress = ""
    
    Dim Change_Report_Up As New Collection
    Dim Change_Report_Down As New Collection
    Const DEV As Boolean = True
    
    With Change_Report
        .Add Change_Report_Up
        .Add Change_Report_Down
    End With

    For Entry = 1 To Entries.Count
        ' Check that the current entry is free of errors
        ' This function will also set the cell value of the current entry according to the Tool Status Page in the WorkBook_Info Collection
        Entry_PreChecks_Passing = Entry_PreChecks(WorkBook_Info, Entries(Entry), Already_Checked_Collection)
        If Not Entry_PreChecks_Passing Then
            GoTo NextEntry
        End If
        
        Compare_and_Update_Color WorkBook_Info, Entries(Entry)
        Comment_State WorkBook_Info, Entries(Entry) 'Insert New Down States into Comments
        Add_Change_to_Report WorkBook_Info, Entries(Entry), Change_Report 'Update Change Report (Currently in main function)
        
NextEntry:
        Progress = PercentDone(CLng(Entry), Entries.Count, "Fly Eagles Fly on the road to victory!")
        Application.StatusBar = "Processing chamber: " & Entry & " / " & Entries.Count & " " & Progress
    Next Entry
    
    Set Process_for_Each_Entry = Change_Report
End Function

Public Function PercentDone(Index As Long, Total As Long, String2Output As String) As String
    Dim ret As String
    ret = ""
    Dim Index_Progress As Double
    Dim String_Progress As Double
    
    Index_Progress = Index / Total
    
    Num_Letters_Shown = Int(Index_Progress * Len(String2Output))
    
    ret = Left(String2Output, Num_Letters_Shown)
    
    PercentDone = ret
End Function


Private Sub Comment_State(ByRef WorkBook_Info As Collection, ByVal Entity_Entry As Entity)
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim CellColor As Long
    CellColor = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(WorkBook_Info(WorkBook_Info.Count)).Interior.Color
    
    Dim Current_Comment As String
    Current_Comment = ""
    If CellColor = 255 Then ' New Red
        Current_Comment = Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(WorkBook_Info(WorkBook_Info.Count)).Row, WorkBook_Info(5)).Value
        
        If Entity_Entry.State = "OutOfControl" Then
           Entity_Entry.State = "OOC"
        End If
        
        If Current_Comment = "" Then
            Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(WorkBook_Info(WorkBook_Info.Count)).Row, WorkBook_Info(5)) = ":" & Entity_Entry.State
        Else
            Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(WorkBook_Info(WorkBook_Info.Count)).Row, WorkBook_Info(5)) = Current_Comment & Chr(10) & ":" & Entity_Entry.State
       End If
    End If
End Sub

Private Sub Add_Change_to_Report(ByRef WorkBook_Info As Collection, ByVal Entity_Entry As Entity, ByRef Change_Report As Collection)
    Const DEV As Boolean = True
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    
    
    Dim Message As String
    Message = ""
    
    Dim CellColor As Long
    CellColor = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(WorkBook_Info(WorkBook_Info.Count)).Interior.Color
    If DEV Then
        If CellColor = 255 Then ' New Red
            Message = Entity_Entry.Name & " went Down: " & Entity_Entry.State
            Change_Report(2).Add Message
        ElseIf CellColor = 5296274 Then 'New Green
            Message = Entity_Entry.Name & " is UTP"
            Change_Report(1).Add Message
        Else
            'Do Nothing
        End If
    Else
        If CellColor = 255 Then ' New Red
            Message = Entity_Entry.Name & " went Down: " & Entity_Entry.State
            Change_Report.Add Message
        ElseIf CellColor = 5296274 Then 'New Green
            Message = Entity_Entry.Name & " is UTP"
            Change_Report.Add Message
        Else
            'Do Nothing
        End If
    End If
End Sub

Private Function Entry_PreChecks(ByRef WorkBook_Info As Collection, ByVal Current_Entity As Entity, ByRef List_of_Checked_Entities As Collection) As Boolean
    Dim ret As Boolean
    ret = False
    
    Dim Legit_Entity As Boolean
    Legit_Entity = Current_Entity.Verify_Entity
    If Legit_Entity = False Then
        Debug.Print ("Not legit: " & Current_Entity.Name)
        Entry_PreChecks_Passing = ret
        Exit Function
    End If
    
    Entity_Already_Checked = Already_Compared_Check(List_of_Checked_Entities, Current_Entity.Name)
    If Entity_Already_Checked = True Then
        Debug.Print ("Already Checked: " & Current_Entity.Name)
        Entry_PreChecks_Passing = ret
        Exit Function
    End If
    
    Entity_Cell_Location = Find_Entity_Cell(WorkBook_Info, Current_Entity.Name)
    If Entity_Cell_Location = "SKIP" Then
        Debug.Print ("Not in Output: " & Current_Entity.Name)
        Entry_PreChecks_Passing = ret
        Exit Function
    End If
    
    ret = True
    Entry_PreChecks = ret
End Function

Private Function Create_Latest_ToolSts_Collection(Input_Option As Long)
    Dim Entries As New Collection
    
    If Input_Option = 1 Then
        'Create Collection from Sheet Values
        Application.StatusBar = "Using Manual Input"
    ElseIf Input_Option = 2 Then
        Application.StatusBar = "Reading file on local drive"
        
        Dim STRING_SQL_OUTPUT_FILE As String
        STRING_SQL_OUTPUT_FILE = "C:\Users\cfarion\AppData\Local\Temp\SQLPathFinder_Temp\out_SQL_Tool_Status.tab"
        
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

    Set Create_Latest_ToolSts_Collection = Entries
End Function

Private Function WorkBook_Setup() As Collection
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const ToolSts_FullHistory_SheetName As String = "ToolStsHistory"
    Const ToolSts_DashBoard_SheetName As String = "Tool Status"
    
    'Prep Output Sheet
    On Error Resume Next
        ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).ShowAllData
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = Global_Functions.Find_Col("Entity", ToolSts_DashBoard_SheetName)
    
    Dim DashBoard_Comments_Column As Long 'Today's comments column
    DashBoard_Comments_Column = Global_Functions.Find_Col("Today's Comments", ToolSts_DashBoard_SheetName)
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(Global_Functions.Column2Letter(DashBoard_Entity_Column) & "65535").End(xlUp).Row
    
    Dim DashBoard_CEID_Column As Long
    DashBoard_CEID_Column = Global_Functions.Find_Col("CEID", ToolSts_DashBoard_SheetName)
    
    Dim DashBoard_Module_Column As Long
    DashBoard_Module_Column = Global_Functions.Find_Col("MODULE", ToolSts_DashBoard_SheetName)

    STRING_Range_Entity_Col = Global_Functions.Column2Letter(DashBoard_Entity_Column) & "1:" & Global_Functions.Column2Letter(DashBoard_Entity_Column) & DashBoard_Last_Row
        
    'This MUST be in alphabetical order or else the algorithms do not work
    ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort.SortFields.Add2 Key:=Range(STRING_Range_Entity_Col), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    With ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName)
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).FreezePanes = True
    
    With ret_Component_Names
        .Add ToolSts_DashBoard_SheetName
        .Add ToolSts_FullHistory_SheetName
        .Add DashBoard_Entity_Column
        .Add DashBoard_Last_Row
        .Add DashBoard_CEID_Column
        .Add DashBoard_Comments_Column
        .Add DashBoard_Module_Column
        .Add TEMP_CELL_HOLDER
    End With
    
    Set WorkBook_Setup = ret_Component_Names
End Function

Private Sub Record_Changes(ByRef WorkBook_Info As Collection, ByRef Changes As Collection)
    Const DEV As Boolean = True
    
    Set Dashboard_WorkBook_Parameters = WorkBook_Info(1)
    Set ToolStsHistory_WorkBook_Parameters = WorkBook_Info(2)
    Set Change_Report_WorkBook_Parameters = WorkBook_Info(3)
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = Dashboard_WorkBook_Parameters(1)
    
    Dim ToolSts_FullHistory_SheetName As String
    ToolSts_FullHistory_SheetName = ToolStsHistory_WorkBook_Parameters(1)
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = Dashboard_WorkBook_Parameters(2)
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = Dashboard_WorkBook_Parameters(3)
    
    Change_Report_SheetName = Change_Report_WorkBook_Parameters(1)
    
    Const Max_Lines_for_MsgBox As Long = 12
    
    Dim History_Last_Column As Long
    History_Last_Column = ToolStsHistory_WorkBook_Parameters(2)
    History_Last_Column = (ActiveWorkbook.Worksheets(ToolSts_FullHistory_SheetName).Range("CLF1").End(xlToLeft).Column + 1)
    
    RANGE_output2paste = Global_Functions.Column2Letter(DashBoard_Entity_Column) & "1:" & Global_Functions.Column2Letter(DashBoard_Entity_Column) & DashBoard_Last_Row
    ActiveWorkbook.Sheets(ToolSts_DashBoard_SheetName).Range(RANGE_output2paste).Copy Destination:=Sheets(ToolSts_FullHistory_SheetName).Range(Global_Functions.Column2Letter(History_Last_Column) & "2")
        
    With ActiveWorkbook.Sheets(ToolSts_FullHistory_SheetName)
        .Range(Global_Functions.Column2Letter(History_Last_Column) & "1").Value = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
        .Columns(History_Last_Column).EntireColumn.AutoFit
    End With
    
    
    If DEV Then
        Num_of_Open = Changes(1).Count
        Num_of_Closed = Changes(2).Count
        
        Open_String = ""
        Closed_String = ""
        
        If Num_of_Open = 0 Then
            Open_String = "No New UTP Entities Since Last Query"
        Else
            For i = 1 To Num_of_Open - 1
                Open_String = Open_String & Changes(1)(i) & Chr(10)
            Next i
            Open_String = Open_String & Changes(1)(i)
        End If
        
        If Num_of_Closed = 0 Then
            Closed_String = "No Down Entities Since Last Query"
        Else
            For i = 1 To Num_of_Closed - 1
                Closed_String = Closed_String & Changes(2)(i) & Chr(10)
            Next i
            Closed_String = Closed_String & Changes(2)(i)
        End If
        
        Time_Format = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
        
        Dim Last_Change_Report_Row As Long
        Last_Change_Report_Row = ActiveWorkbook.Sheets("Change Report").Range("A1812").End(xlUp).Row + 1
        
        With Sheets(Change_Report_SheetName)
            .Cells(Last_Change_Report_Row, 1) = Time_Format
            .Cells(Last_Change_Report_Row, 2) = Open_String
            .Cells(Last_Change_Report_Row, 3) = Closed_String
        End With
        If (Num_of_Open + Num_of_Closed) > Max_Lines_for_MsgBox Then
            Change_String = "Multiple Changes above Message Box Limit. Please check the Change Report Tab."
        Else
            Change_String = "CHANGE_REPORT: " & Chr(10) & Chr(10) & Open_String & Chr(10) & Closed_String
        End If
        msg = MsgBox(Change_String, 0, Time_Format, 0, 0)
    Else
        Dim Num_of_Changes As Long
        Num_of_Changes = Changes.Count
        
        'Dim Change_String As String
        Change_String = ""
        
        If Num_of_Changes <> 0 Then
            For Current_Change = 1 To Num_of_Changes - 1
                Change_String = Change_String & Changes(Current_Change) & Chr(10)
            Next Current_Change
            Change_String = Change_String & Changes(Current_Change) 'Eliminate the last carriage return
        Else
            Change_String = "No Changes"
        End If
        
        Time_Format = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
        
        'Dim Last_Change_Report_Row As Long
        Last_Change_Report_Row = ActiveWorkbook.Sheets("Change Report").Range("A1812").End(xlUp).Row + 1
        
        With Sheets(Change_Report_SheetName)
            .Cells(Last_Change_Report_Row, 1) = Time_Format
            .Cells(Last_Change_Report_Row, 2) = Change_String
        End With
        If Num_of_Changes > Max_Lines_for_MsgBox Then
            Change_String = "Multiple Changes above Message Box Limit. Please check the Change Report Tab."
        Else
            Change_String = "CHANGE_REPORT: " & Chr(10) & Chr(10) & Change_String
        End If
        msg = MsgBox(Change_String, 0, Time_Format, 0, 0)
    End If
End Sub

Private Sub Populate_Entry_Collection(ByRef coll As Collection, ByVal File_Text As String)
    Dim Entry As Entity
    
    Dim Name_index As Long
    Dim Availability_index As Long
    Dim State_index As Long
    Dim CEID_index As Long
    
    Instances = Split(File_Text, Chr(10))
    For Instance = 0 To UBound(Instances) - 1
        Set Entry = New Entity
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
                Else
                    'Error / Do Nothing
                End If
            Next i
        Else
            If UBound(Components) > 0 Then
                Entry.Name = Components(Name_index)
                Entry.Availability = Components(Availability_index)
                Entry.State = Components(State_index)
                Entry.CEID = Components(CEID_index)
                coll.Add Entry
            End If
        End If
    Next Instance
End Sub

Private Sub SQL_Script_Through_Excel(ByRef coll As Collection)
    On Error GoTo ErrHandler

    Dim DataSource As String
    DataSource = "D1D_PROD_XEUS"
    
    Dim asd As String

    Dim sQuery As String
    sQuery = ActiveWorkbook.Sheets("SQL_INPUT").Cells(2, 2).Value

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
    Dim Entry As Entity
    Dim field As Object
    
    ' Loop through the recordset and add each row to the collection
    Do While Not recordSet.EOF
        Set Entry = New Entity
        For Each field In recordSet.Fields
            fieldName = field.Name
            If fieldName = "TOOL_NAME" Then
                Entry.Name = field.Value
            ElseIf fieldName = "AVAILABILITY" Then
                Entry.Availability = field.Value
            ElseIf fieldName = "STATE" Then
                Entry.State = field.Value
            ElseIf fieldName = "CEID" Then
                Entry.CEID = field.Value
            Else
                'Error / Do Nothing
            End If
        Next field
        coll.Add Entry
        recordSet.MoveNext
        asd = CStr(Entry.Name & ": " & Entry.Availability)
        If Entry.Availability = "Down" Then
            asd = asd & " -> " & CStr(Entry.State)
        End If
        LogCollection.Add asd
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

'TODO: Change back to Private after Testing
Private Function Find_Entity_Cell(ByRef WorkBook_Info As Collection, ByVal Entity As String) As String
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
        If current_value = Entity Then
            first_row = DashBoard_Last_Row + 1
        ElseIf current_value < Entity Then
            first_row = Current_Row
        ElseIf current_value > Entity Then
            DashBoard_Last_Row = Current_Row
        End If
    Wend
    
    If (current_value = Entity) Then
        ret = Global_Functions.Column2Letter(DashBoard_Entity_Column) & Current_Row
    ElseIf (Cells(Current_Row + 1, DashBoard_Entity_Column) = Entity) Then
        ret = Global_Functions.Column2Letter(DashBoard_Entity_Column) & Current_Row + 1
    Else
        'Not in Tool Status Page
    End If
    WorkBook_Info.Remove WorkBook_Info.Count
    WorkBook_Info.Add ret
    Find_Entity_Cell = ret
End Function

Private Function Already_Compared_Check(ByRef Compared_Collection As Collection, ByVal Entity As String) As Boolean
    Dim ret As Boolean
    ret = False
    
    Dim lower_bound As Long
    Dim upper_bound As Long
    
    'Empty collection case
    If Compared_Collection.Count = 0 Then
        Compared_Collection.Add item:=Entity
        Already_Compared_Check = ret
        Exit Function
    End If
    
    'Binary Search parameters
    lower_bound = 1
    upper_bound = Compared_Collection.Count
    
    While (lower_bound <> upper_bound - 1) And (lower_bound <> upper_bound)
        current_index = Int((lower_bound + upper_bound) \ 2)
        'current_index = Application.WorksheetFunction.Ceiling((lower_bound + upper_bound) \ 2, 1)
        current_index_entity = Compared_Collection(current_index)
        
        If current_index_entity = Entity Then
            ret = True
            Already_Compared_Check = ret
            Exit Function
        ElseIf current_index_entity < Entity Then
            lower_bound = current_index
        Else
            upper_bound = current_index
        End If
    Wend
    
    If Compared_Collection(lower_bound) = Entity Or Compared_Collection(upper_bound) = Entity Then
        ret = True
        Already_Compared_Check = ret
        Exit Function
    End If
    
    check1 = Compared_Collection(lower_bound) < Entity
    check2 = Compared_Collection(upper_bound) < Entity
    
    If check1 = False And check2 = False Then
        'Before is equivalent to replace the lower_bound index
        Compared_Collection.Add item:=Entity, Before:=lower_bound
    ElseIf check1 = True And check2 = False Then
        'Before is equivalent to replace the upper_bound index
        Compared_Collection.Add item:=Entity, Before:=upper_bound
    ElseIf check1 = True And check2 = True Then
        'Placing the item at the end of the list
        Compared_Collection.Add item:=Entity, After:=upper_bound
    Else
        'Error
    End If
    
    Already_Compared_Check = ret
End Function

Private Sub Compare_and_Update_Color(ByRef WorkBook_Info As Collection, ByVal Entity_Entry As Entity)
    Dim NewGreen As Long
    Dim OldGreen As Long
    Dim NewRed As Long
    Dim OldRed As Long
    Dim TextBlack As Long
    Dim TextWhite As Long
    
    NewGreen = 5296274  '#92D050
    OldGreen = 9498256  '#90EE90
    NewRed = 255        '#FF0000
    OldRed = 8421616    '#F08080
    TextBlack = 0       '#000000
    TextWhite = 16777215 '#FFFFFF
    
    
    Dim Entity_Availability As String
    Entity_Availability = Entity_Entry.Availability
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim TEMP_CELL_HOLDER As String
    TEMP_CELL_HOLDER = WorkBook_Info(WorkBook_Info.Count)

    Dim CurrentCellColor As Long
    CurrentCellColor = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(TEMP_CELL_HOLDER).Cells.Interior.Color
    
    Dim CurrentCellTextColor As Long
    CurrentCellTextColor = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(TEMP_CELL_HOLDER).Font.Color
    
    Dim NewCellColor As Long
    NewCellColor = CurrentCellColor 'If all else fails, leave it as its current color
    
    Dim NewCellTextColor As Long
    NewCellTextColor = CurrentCellTextColor 'If all else fails, leave the text as its current color
    
    If CurrentCellColor = OldGreen And Entity_Availability = "Down" Then
        NewCellColor = NewRed
        NewCellTextColor = TextWhite
    ElseIf CurrentCellColor = OldRed And Entity_Availability = "Up" Then
        NewCellColor = NewGreen
        NewCellTextColor = TextWhite
    ElseIf CurrentCellColor = NewGreen Then
        If Entity_Availability = "Down" Then
            NewCellColor = NewRed
            NewCellTextColor = TextWhite
        ElseIf Entity_Availability = "Up" Then
            NewCellColor = OldGreen
            NewCellTextColor = TextBlack
        Else
            'Error
        End If
    ElseIf CurrentCellColor = NewRed Then
        If Entity_Availability = "Down" Then
            NewCellColor = OldRed
            NewCellTextColor = TextBlack
        ElseIf Entity_Availability = "Up" Then
            NewCellColor = NewGreen
            NewCellTextColor = TextWhite
        Else
            'Error
        End If
    Else
        'Error or cell not one of the four colors
    End If
    
    With ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Range(TEMP_CELL_HOLDER)
        .Cells.Interior.Color = NewCellColor
        .Font.Color = NewCellTextColor
    End With
End Sub

Public Function STRING_get_last_row_and_col(Optional SHEET_NAME As String, Optional STARTING_ROW As Long, Optional STARTING_COL As Long) As String
    Dim sheet2use As String
    Dim row2use As Long
    row2use = 1
    Dim col2use As Long
    col2use = 1
    
    If SHEET_NAME = "" Then
        sheet2use = ActiveSheet.Name
    ElseIf SHEET_NAME <> "" Then
        sheet2use = ActiveWorkbook.Worksheets(SHEET_NAME).Name
    End If
    
    If STARTING_ROW <> 0 Then
        row2use = STARTING_ROW
    End If
    
    If STARTING_COL <> 0 Then
        col2use = STARTING_COL
    End If
    
    'Find last row and column
    'Last_Row = ActiveWorkbook.Worksheets(sheet2use).Range("A65535").End(xlUp).Row
    Last_Row = ActiveWorkbook.Worksheets(sheet2use).Cells(65535, col2use).End(xlUp).Row
    
    'last_col = ActiveWorkbook.Worksheets(sheet2use).Range("CLF1").End(xlToLeft).Column
    last_col = ActiveWorkbook.Worksheets(sheet2use).Cells(row2use, 215).End(xlToLeft).Column
    STRING_get_last_row_and_col = Last_Row & ";" & last_col
End Function

Sub InitializeLogCollection()
    Set LogCollection = New Collection
End Sub

Sub InitializeCommentCollection()
    Set CommentCollection = New Collection
    
    Dim timestamp As String
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    CommentCollection.Add timestamp
End Sub

Sub SaveLogsToTextFile()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim item As Variant
    Dim timestamp As String
    
    ' Generate a timestamp for the file name
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    
    ' Set the file path where you want to save the text file
    filePath = "C:\Users\cfarion\OneDrive - Intel Corporation\00_Work_OneDrive\03_Personal_Scripts\_Logs\" & timestamp & "_ToolSts.txt"
    
    ' Get the next available file number
    fileNumber = FreeFile
    
    ' Open the text file for output
    Open filePath For Output As #fileNumber
    
    ' Loop through each item in the Collection and write to the text file
    For Each item In LogCollection
        Print #fileNumber, item
    Next item
    
    ' Close the text file
    Close #fileNumber
    
    ' Notify the user
    'MsgBox "The strings have been saved to the text file: " & filePath
End Sub

Sub SaveCommentsToTextFile()
    Dim filePath As String
    Dim fileNumber As Integer
    Dim item As Variant
    Dim timestamp As String
    
    ' Generate a timestamp for the file name
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    
    ' Set the file path where you want to save the text file
    filePath = "C:\Users\cfarion\OneDrive - Intel Corporation\00_Work_OneDrive\03_Personal_Scripts\_Logs\" & timestamp & "_Comments.txt"
    
    ' Get the next available file number
    fileNumber = FreeFile
    
    ' Open the text file for output
    Open filePath For Output As #fileNumber
    
    ' Loop through each item in the Collection and write to the text file
    For Each item In CommentCollection
        Print #fileNumber, item
    Next item
    
    ' Close the text file
    Close #fileNumber
    
    ' Notify the user
    'MsgBox "The strings have been saved to the text file: " & filePath
End Sub

