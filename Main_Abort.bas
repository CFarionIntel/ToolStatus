Attribute VB_Name = "Main_Abort"
Sub Main_Abort_Processing()
    Debug.Print ("Start Main > " & Format(Time, "hh:mm.ss"))
    'Collect the Sheet Names and columns which comprise the Dashboard
    Dim WorkBook_Input_Parameters As Collection
    Set WorkBook_Input_Parameters = Abort_Input_WorkBook_Setup
    
    'Select how you want to collect the new Tool Status Update
    Dim Abort_Input_Selection As Long
    Abort_Input_Selection = ActiveWorkbook.Sheets("Settings").Cells(1, 7)
    
    'Create a collection of the latest Tool Status from the Input
    Dim Entries As Collection
    Set Entries = Create_Latest_Abort_Collection(WorkBook_Input_Parameters, Abort_Input_Selection)
    
    If Entries.Count <> 0 Then
        Dim WorkBook_Output_Parameters As Collection
        Set WorkBook_Output_Parameters = Abort_History_WorkBook_Setup
    
        'Dispo each wafer
        Dispo_Wafers Entries
        
        'Display Individual Wafers on Input Sheet
        Record_Dispos Entries, WorkBook_Input_Parameters
        
        Get_Chambers_for_E3_Search WorkBook_Input_Parameters, Entries
        
        'What is typically displayed from a SQL script
        Record_Query Entries, WorkBook_Output_Parameters
        
        Display_Abort_Statistics Entries, WorkBook_Output_Parameters
    End If
    
    Application.StatusBar = False
    Debug.Print ("End Main > " & Format(Time, "hh:mm.ss"))
End Sub

Sub Display_Abort_Statistics(ByRef Entries As Collection, ByRef WorkBook_Info As Collection)
    Dim Entry_Statistics As Collection
    Dim Process_Time As Double
    Dim Sum_of_Process As Double
    Dim Non_MMO As Long
    Non_MMO = 0
    
    Abort_Setup_Sheetname = WorkBook_Info(1)
    
    Dim Time_Diff_Col As Long
    Time_Diff_Col = WorkBook_Info(11)
    Sum_of_Process = 0
    
    If Entries.Count = 0 Then
        Exit Sub
    End If
    
    For i = 1 To Entries.Count
        If Entries(i).Dispo <> "RMI" Then
            Process_Time = Entries(i).ProcessTime
            Sum_of_Process = Sum_of_Process + Process_Time
            ActiveWorkbook.Sheets(Abort_Setup_Sheetname).Cells(i + 1, Time_Diff_Col) = Round(Process_Time, 3)
        Else
            Non_MMO = Non_MMO + 1
        End If
    Next i
    
    If (Entries.Count - Non_MMO) <> 0 Then
        Average_Process_Time = Sum_of_Process / (Entries.Count - Non_MMO)
        ActiveWorkbook.Sheets(Abort_Setup_Sheetname).Cells(2, Time_Diff_Col + 1) = Round(Average_Process_Time, 3)
        
        Sum_of_Process = 0
        For i = 1 To Entries.Count
            If Entries(i).Dispo <> "RMI" Then
                Process_Time = Entries(i).ProcessTime
                Diff_sq = (Process_Time - Average_Process_Time) ^ 2
                Sum_of_Process = Sum_of_Process + Diff_sq
                ActiveWorkbook.Sheets(Abort_Setup_Sheetname).Cells(i + 1, Time_Diff_Col + 2) = Round(Diff_sq, 3)
            End If
        Next i
        
        sigma = (Sum_of_Process / (Entries.Count - Non_MMO)) ^ (1 / 2)
        ActiveWorkbook.Sheets(Abort_Setup_Sheetname).Cells(2, Time_Diff_Col + 3) = Round(sigma, 3)
    End If
    
End Sub

Public Sub Record_Dispos(ByRef Entries As Collection, ByRef WorkBook_Info As Collection)
    Dim Abort_Input_SheetName As String
    Abort_Input_SheetName = WorkBook_Info(1)
    
    Dim Number_of_MMO As Long
    Number_of_MMO = 0
    Dim Number_of_RMI As Long
    Number_of_RMI = 0
    Dim Number_of_Partial As Long
    Number_of_Partial = 0
    Dim Number_of_Error As Long
    Number_of_Error = 0
    
    Abort_Output_Column = WorkBook_Info(10)
    Lot_Output_Row = WorkBook_Info(11)
    Operation_Output_Row = WorkBook_Info(12)
    Entity_Output_Row = WorkBook_Info(13)
    
    MMO_Output_Row = WorkBook_Info(14)
    RMI_Output_Row = WorkBook_Info(15)
    Partial_Output_Row = WorkBook_Info(16)
    Error_Output_Row = WorkBook_Info(17)
    
    
    'Change Range Below to be non hard coded
    ActiveWorkbook.Sheets(Abort_Input_SheetName).Range("D11:AB18").ClearContents
    
    Dim Wafer_Name As String
    
    ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Lot_Output_Row, Abort_Output_Column) = UCase(Entries(1).Lot)
    ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Operation_Output_Row, Abort_Output_Column) = Entries(1).Operation
    ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Entity_Output_Row, Abort_Output_Column) = UCase(Entries(1).Entity)
    
    For i = 1 To Entries.Count
        Wafer_Name = "[" & Entries(i).Slot & "/" & Format(Entries(i).Waf3, "000") & "]"
        
        Select Case Entries(i).Dispo
            Case "MMO"
                ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(MMO_Output_Row, Abort_Output_Column + Number_of_MMO) = Wafer_Name
                Number_of_MMO = Number_of_MMO + 1
            Case "RMI"
                ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(RMI_Output_Row, Abort_Output_Column + Number_of_RMI) = Wafer_Name
                Number_of_RMI = Number_of_RMI + 1
            Case "RMI Full Ash SIF"
            Case "Partial"
                ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Partial_Output_Row, Abort_Output_Column + Number_of_Partial) = Wafer_Name
                Number_of_Partial = Number_of_Partial + 1
            Case "Error"
                ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Error_Output_Row, Abort_Output_Column + Number_of_Error) = Wafer_Name
                Number_of_Error = Number_of_Error + 1
            Case Else
                'Do Nothing
        End Select
    Next i
    
End Sub

Public Sub Dispo_Wafers(ByRef Entries As Collection)
    'Entry into this sub already checked if there are no wafers to dispo
    Const DEV As Boolean = True
    
    Dim Number_of_Wafers As Long
    Number_of_Wafers = Entries.Count
    
    Dim Ashing_Required As Boolean
    Dim Wafer_Etched As Boolean
    Dim Wafer_Ashed As Boolean
    Dim Etched_and_Ashed_Result As String
    Dim Temp_Dispo_Result As String
    
    
    ret = "Error"
    
    MMO_sum = 0
    MMO_Samples = 0
    
    'For each wafer
    ''Determine if ashing is required
    ''Did it enter all of the appropriate chambers?
    For i = 1 To Entries.Count
        Ashing_Required = False
        Wafer_Etched = False
        Wafer_Ashed = False
        Etched_and_Ashed_Result = Wafer_Etched & ";" & Wafer_Ashed
        If InStr(1, Entries(i).Recipe, "HYBRID") <> 0 Or InStr(1, Entries(i).Recipe, "LK") Then
            Ashing_Required = True
        End If
        
        Path_Steps = Split(Entries(i).ChamberPath, ";")
        Etched_and_Ashed_Result = Wafer_Chamber_Entry(Path_Steps)
        Entries(i).Dispo = Determine_Dispo(Etched_and_Ashed_Result, Ashing_Required)
    Next i
    
    Determine_Partials Entries
End Sub

Private Sub Determine_Partials(ByRef Entries As Collection)
    If Entries(1).DateStart <> "No Start Date" Then
        Dim mu As Double
        Dim sigma As Double
        Dim average_sum As Double
        Dim variance_sum As Double
        Dim sqr_error As Double
        Dim MMO_Samples As Long
        Dim Process_Times As Collection
        Dim Boundary_Result As Double
        
        Const e As Double = 2.71828182845904
        Const Partial_Determinate_Value As Double = 0.9 'Determines when a wafer is considered a partial
        
        average_sum = 0
        variance_sum = 0
        sqr_error = 0
    
        For i = 1 To Entries.Count
            If Entries(i).Dispo = "MMO" Then
                Time_Diff = Entries(i).ProcessTime
                average_sum = average_sum + Time_Diff
                MMO_Samples = MMO_Samples + 1
            End If
        Next i
        
        If MMO_Samples <> 0 Then
            mu = average_sum / MMO_Samples
            
            For i = 1 To Entries.Count
                If Entries(i).Dispo = "MMO" Then
                    Time_Diff = Entries(i).ProcessTime
                    sqr_error = (Time_Diff - mu) ^ 2
                    
                    variance_sum = variance_sum + sqr_error
                End If
            Next i
                    
            sigma = (variance_sum / MMO_Samples) ^ (1 / 2)
            
            For i = 1 To Entries.Count
                If Entries(i).Dispo = "MMO" Then
                    Time_Diff = Entries(i).ProcessTime
                    
                    denominator = 1 + e ^ -(Time_Diff - (mu + sigma))
                    Boundary_Result = 1 / denominator
                    If Boundary_Result > Partial_Determinate_Value Then
                        Entries(i).Dispo = "Partial"
                    End If
                End If
            Next i
        End If
    End If
End Sub

Private Function Determine_Dispo(Etch_and_Ashed_Result As String, Ashing_Required As Boolean)
    Dim ret As String
    ret = "Error"
    Dim Etched_Result As Boolean
    Dim Ashed_Result As Boolean
    
    Etched_Result = CBool(Split(Etch_and_Ashed_Result, ";")(0))
    Ashed_Result = CBool(Split(Etch_and_Ashed_Result, ";")(1))
    
    If Etched_Result Then
        ret = "MMO"
    ElseIf Not Etched_Result Then
        ret = "RMI"
    Else
        'Do Nothing / Error
    End If
    
    If Ashing_Required Then
        If Etched_Result And Ashed_Result Then
            'Already covered in case above as "MMO"
        ElseIf Etched_Result And Not Ashed_Result Then
            ret = "RMI Full Ash SIF"
        Else
            'Do Nothing / Error
        End If
    End If
    Determine_Dispo = ret
End Function

Private Function Wafer_Chamber_Entry(Path As Variant) As String
    Dim Wafer_Etched As Boolean
    Dim Wafer_Ashed As Boolean
    Dim First_Char As String
    
    Wafer_Etched = False
    Wafer_Ashed = False
    
    If UBound(Path) <> 0 Then
        Dim Current_Step As Integer
        Current_Step = 0
        While (Current_Step <= UBound(Path)) 'Add condition to exit loop if etching already determined
            First_Char = Left(Path(Current_Step), 1)
            If First_Char = "P" Then
                If Right(Path(Current_Step), 1) < 7 Then
                    Wafer_Etched = True
                ElseIf Right(Path(Current_Step), 1) >= 7 Then
                    Wafer_Ashed = True
                Else
                    'Do Nothing
                End If
            End If
            Current_Step = Current_Step + 1
        Wend
    End If
    
    Wafer_Chamber_Entry = Wafer_Etched & ";" & Wafer_Ashed
End Function

Public Sub Record_Query(ByRef Entries As Collection, ByRef WorkBook_Info As Collection)
    Time_Format = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
    Const DEV As Boolean = True
    
    Dim Abort_FullHistory_SheetName As String
    Abort_FullHistory_SheetName = WorkBook_Info(1)
    
    Dim Number_of_Wafers As Long
    Number_of_Wafers = Entries.Count
    
    Dim Range_for_Inserting_Rows As String
    Range_for_Inserting_Rows = CStr(2) & ":" & CStr(2 + Number_of_Wafers - 1)
    ActiveWorkbook.Sheets(Abort_FullHistory_SheetName).Range(Range_for_Inserting_Rows).EntireRow.Insert
    With ActiveWorkbook.Sheets(Abort_FullHistory_SheetName).Range(Range_for_Inserting_Rows).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With

    For i = 1 To (Number_of_Wafers)
        If DEV Then
            With ActiveWorkbook.Sheets(Abort_FullHistory_SheetName)
            .Cells(i + 1, 1) = Time_Format
            .Cells(i + 1, 2) = Entries(i).Entity
            .Cells(i + 1, 3) = Entries(i).Lot
            .Cells(i + 1, 4) = Entries(i).Operation
            .Cells(i + 1, 5) = Entries(i).Slot
            .Cells(i + 1, 6) = Format(CLng(Entries(i).Waf3), "000")
            .Cells(i + 1, 7) = Entries(i).ChamberPath
            .Cells(i + 1, 8) = Entries(i).Recipe
            .Cells(i + 1, 9) = Entries(i).DateStart
            .Cells(i + 1, 10) = Entries(i).DateEnd
            .Cells(i + 1, 11) = Entries(i).ProcessTime
        End With
        Else
        With ActiveWorkbook.Sheets(Abort_FullHistory_SheetName)
            .Cells(i + 1, 1) = Time_Format
            .Cells(i + 1, WorkBook_Info(3)) = Entries(i).Entity
            .Cells(i + 1, WorkBook_Info(4)) = Entries(i).Lot
            .Cells(i + 1, WorkBook_Info(5)) = Entries(i).Operation
            .Cells(i + 1, WorkBook_Info(6)) = Entries(i).Slot
            .Cells(i + 1, WorkBook_Info(7)) = Format(CLng(Entries(i).Waf3), "000")
            .Cells(i + 1, WorkBook_Info(8)) = Entries(i).ChamberPath
            .Cells(i + 1, WorkBook_Info(9)) = Entries(i).Recipe
            .Cells(i + 1, WorkBook_Info(10)) = Entries(i).DateStart
            .Cells(i + 1, WorkBook_Info(11)) = Entries(i).DateEnd
            .Cells(i + 1, WorkBook_Info(12)) = Entries(i).ProcessTime
        End With
        End If
    Next i
End Sub

Private Function Create_Latest_Abort_Collection(ByRef WorkBook_Info As Collection, ByVal Input_Option As String)
    Dim Entries As New Collection
    
    If Input_Option = 1 Then
        'Create Collection from Sheet Values
    ElseIf Input_Option = 2 Then
        Dim STRING_SQL_OUTPUT_FILE As String
        STRING_SQL_OUTPUT_FILE = "C:\Users\cfarion\AppData\Local\Temp\SQLPathFinder_Temp\out_24452.tab"
        
        Dim Text As String
        Text = Global_Functions.Read_SQL_Output_File(STRING_SQL_OUTPUT_FILE)
        
        Populate_Entry_Collection Entries, Text
    ElseIf Input_Option = 3 Then
        SQL_Script_Through_Excel Entries, WorkBook_Info
    Else
        'Nothing Selected or Error
    End If

    Set Create_Latest_Abort_Collection = Entries
End Function

Private Function WorkBook_Setup() As Collection
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const Abort_Input_SheetName As String = "Abort Setup"
    Const Abort_FullHistory_SheetName As String = "AbortHistory"
    
    Dim Abort_History_Entity_Column As Long
    Abort_History_Entity_Column = Global_Functions.Find_Col("ENTITY", Abort_FullHistory_SheetName)
    
    Dim Abort_History_Lot_Column As Long
    Abort_History_Lot_Column = Global_Functions.Find_Col("LOT", Abort_FullHistory_SheetName)
    
    Dim Abort_History_Operation_Column As Long
    Abort_History_Operation_Column = Global_Functions.Find_Col("OPERATION", Abort_FullHistory_SheetName)
    
    Dim Abort_History_Slot_Column As Long
    Abort_History_Slot_Column = Global_Functions.Find_Col("SLOT", Abort_FullHistory_SheetName)
    
    Dim Abort_History_WAF3_Column As Long
    Abort_History_WAF3_Column = Global_Functions.Find_Col("WAF3", Abort_FullHistory_SheetName)
    
    Dim Abort_History_ChamberPath_Column As Long
    Abort_History_ChamberPath_Column = Global_Functions.Find_Col("CHAMBER_PATH", Abort_FullHistory_SheetName)
    
    Dim Abort_History_Recipe_Column As Long
    Abort_History_Recipe_Column = Global_Functions.Find_Col("RECIPE", Abort_FullHistory_SheetName)
    
    Dim Abort_History_Start_Column As Long
    Abort_History_Start_Column = Global_Functions.Find_Col("WAFER_ENTITY_START_DATE", Abort_FullHistory_SheetName)
    
    Dim Abort_History_End_Column As Long
    Abort_History_End_Column = Global_Functions.Find_Col("WAFER_ENTITY_END_DATE", Abort_FullHistory_SheetName)
    
    Dim Abort_History_ProcessTime_Column As Long
    Abort_History_ProcessTime_Column = Global_Functions.Find_Col("CHAMBER_PROCESS_DURATION", Abort_FullHistory_SheetName)
        
    Dim Abort_History_Process_Time_Column As Long
    Abort_History_Process_Time_Column = Global_Functions.Find_Col("Process Time", Abort_FullHistory_SheetName)
        
    With ret_Component_Names
        .Add Abort_Input_SheetName
        .Add Abort_FullHistory_SheetName
        .Add Abort_History_Entity_Column
        .Add Abort_History_Lot_Column
        .Add Abort_History_Operation_Column
        .Add Abort_History_Slot_Column
        .Add Abort_History_WAF3_Column
        .Add Abort_History_ChamberPath_Column
        .Add Abort_History_Recipe_Column
        .Add Abort_History_Start_Column
        .Add Abort_History_End_Column
        .Add Abort_History_ProcessTime_Column
        .Add TEMP_CELL_HOLDER
    End With
    
    Set WorkBook_Setup = ret_Component_Names
End Function

Private Sub Populate_Entry_Collection(ByRef coll As Collection, ByVal File_Text As String)
    Dim Entry As Abort
    
    Dim Name_index As Long
    Dim Lot_index As Long
    Dim Operation_index As Long
    Dim Slot_index As Long
    Dim Waf3_index As Long
    Dim ChamberPath_index As Long
    Dim Recipe_index As Long
    Dim DateStart_index As Long
    Dim DateEnd_index As Long
    Dim ProcessTime_index As Long
    
    Instances = Split(File_Text, Chr(10))
    For Instance = 0 To UBound(Instances) - 1
        Set Entry = New Abort
        Components = (Split(Instances(Instance), Chr(9)))
        If Instance = 0 Then 'Determine which column is which variable. Good for if someone changes the output columns of SQL script
            For i = 0 To UBound(Components)
                If Components(i) = "ENTITY" Then
                    Name_index = i
                ElseIf Components(i) = "LOT" Then
                    Lot_index = i
                ElseIf Components(i) = "OPERATION" Then
                    Operation_index = i
                ElseIf Components(i) = "SLOT" Then
                    Slot_index = i
                ElseIf Components(i) = "WAF3" Then
                    Waf3_index = i
                ElseIf Components(i) = "CHAMBER_PATH" Then
                    ChamberPath_index = i
                ElseIf Components(i) = "RECIPE" Then
                    Recipe_index = i
                ElseIf Components(i) = "WAFER_ENTITY_START_DATE" Then
                    DateStart_index = i
                ElseIf Components(i) = "WAFER_ENTITY_END_DATE" Then
                    DateEnd_index = i
                ElseIf Components(i) = "CHAMBER_PROCESS_DURATION" Then
                    ProcessTime_index = i
                Else
                    'Error / Do Nothing
                End If
            Next i
        Else
            If UBound(Components) > 0 Then
                With Entry
                    .Entity = Components(Name_index)
                    .Lot = Components(Lot_index)
                    On Error Resume Next
                        .Operation = CLng(Components(Operation_index))
                    On Error Resume Next
                        .Slot = CLng(Components(Slot_index))
                    On Error Resume Next
                        .Waf3 = CLng(Components(Waf3_index))
                    .ChamberPath = Components(ChamberPath_index)
                    .Recipe = Components(Recipe_index)
                    .DateStart = Components(DateStart_index)
                    .DateEnd = Components(DateEnd_index)
                    On Error Resume Next
                        .ProcessTime = CDbl(Components(ProcessTime_index))
                End With
                
                coll.Add Entry
            End If
        End If
    Next Instance
End Sub

Private Sub SQL_Script_Through_Excel(ByRef coll As Collection, ByRef WorkBook_Info As Collection)
    On Error GoTo ErrHandler
    Const DEV As Boolean = True
        
    Dim Abort_Input_SheetName As String
    Abort_Input_SheetName = WorkBook_Info(1)

    Dim DataSource As String
    DataSource = "D1D_PROD_XEUS"

    Dim sQuery As String
    
    Dim Input_Query As String
    Input_Query = ActiveWorkbook.Sheets("SQL_INPUT").Cells(9, 2).Value
    
    If DEV Then
        Input_Query = ActiveWorkbook.Sheets("SQL_INPUT").Cells(10, 2).Value
    End If
    
    sQuery = Build_Abort_Query(WorkBook_Info, Input_Query)
    
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

    Dim Entry As Abort
    
    ' Check if the recordset is empty
    If recordSet.EOF And recordSet.BOF Then
        'MsgBox "No records found."
        Set Entry = New Abort
        Entry.Entity = UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(4, 2).Value)
        Entry.Lot = ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(1, 2).Value
        Entry.Operation = ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(2, 2).Value
        Entry.Slot = 0
        Entry.Waf3 = 0
        Entry.ChamberPath = "Never entered Entity"
        Entry.Recipe = "Please find through other means"
        Entry.DateStart = "No Start Date"
        Entry.DateEnd = "No End Date"
        Entry.ProcessTime = 0
        coll.Add Entry
        Exit Sub
    End If

    ' Initialize the collection
    Dim field As Object
    
    ' Loop through the recordset and add each row to the collection
    Do While Not recordSet.EOF
        Set Entry = New Abort
        For Each field In recordSet.Fields
            fieldName = field.Name
            If fieldName = "ENTITY" Then
                Entry.Entity = field.Value
            ElseIf fieldName = "LOT" Then
                Entry.Lot = field.Value
            ElseIf fieldName = "OPERATION" Then
                Entry.Operation = field.Value
            ElseIf fieldName = "SLOT" Then
                Entry.Slot = field.Value
            ElseIf fieldName = "WAF3" Then
                Entry.Waf3 = field.Value
            ElseIf fieldName = "CHAMBER_PATH" Then
                Entry.ChamberPath = field.Value
            ElseIf fieldName = "RECIPE" Then
                Entry.Recipe = field.Value
            ElseIf fieldName = "WAFER_ENTITY_START_DATE" Then
                Entry.DateStart = field.Value
            ElseIf fieldName = "WAFER_ENTITY_END_DATE" Then
                Entry.DateEnd = field.Value
            ElseIf fieldName = "CHAMBER_PROCESS_DURATION" Then
                Entry.ProcessTime = field.Value
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

Private Function Build_Abort_Query(ByRef WorkBook_Info As Collection, ByVal iQuery As String) As String
    Dim ret_Finished_Query As String
    ret_Finished_Query = ""
    
    Dim Abort_Input_SheetName As String
    Abort_Input_SheetName = WorkBook_Info(1)
    
    Dim Input_Values As New Collection

    Dim Abort_Input_Column As Long
    Dim Lot_Input_Row As Long
    Dim Operation_Input_Row As Long
    Dim Entity_Input_Row As Long
    Dim Error_Message_Row As Long
    Dim DaysBack_Input_Row As Long
    
    Abort_Input_Column = WorkBook_Info(2)
    Lot_Input_Row = WorkBook_Info(3)
    Operation_Input_Row = WorkBook_Info(4)
    Entity_Input_Row = WorkBook_Info(6)
    Error_Message_Row = WorkBook_Info(8)
    DaysBack_Input_Row = WorkBook_Info(9)
    
    Dim Error_Message As String
    Error_Message = ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Error_Message_Row, Abort_Input_Column).Value
    
    If Error_Message <> "" Then
        Dim Error_Msg_Components As New Collection
        Set Error_Msg_Components = Parse_Error_Msg(Error_Message)
    End If
        
    With Input_Values
        .Add Clean_Up_Input(UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Lot_Input_Row, Abort_Input_Column).Value))
        .Add Clean_Up_Input(UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Entity_Input_Row, Abort_Input_Column).Value))
        .Add Clean_Up_Input(UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(Operation_Input_Row, Abort_Input_Column).Value))
        .Add Clean_Up_Input(UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(DaysBack_Input_Row, Abort_Input_Column).Value))
    End With
    
    If Input_Values(4) = "" Then
       Input_Values.Remove 4
       Input_Values.Add 2
    End If
    
    Query_Components = Split(iQuery, "..::..")
    
    Dim Number_of_Components As Long
    Number_of_Components = UBound(Query_Components)
    ret_Finished_Query = ret_Finished_Query & Query_Components(0)
    
    For i = 1 To Number_of_Components
        ret_Finished_Query = ret_Finished_Query & Input_Values(i) & Query_Components(i)
    Next i
        
    Build_Abort_Query = ret_Finished_Query
End Function

Private Function BreakUp_Error_Msg(ByRef WorkBook_Info As Collection, ByVal Error_Message As String) As Collection
    'This assumes E3 message format has not changed
    Dim ret As New Collection
    
    Dim E3_Clue_String As String
    E3_Clue_String = "FDC excursion"
    
    Dim E3_Error_Type As Long
    E3_Error_Type = 0
    
    Dim Error_Message_Entity As String
        
    E3_Error_Type = InStr(1, Error_Message, E3_Clue_String)
    If E3_Error_Type <> 0 Then
        Error_Message_Components = Split(Error_Message, ":")
        Error_Message_Lot = Split(Error_Message_Components(1), ",")(0)
        Error_Message_Entity = Split(Error_Message_Components(2), ",")(0)
    
        Cell_Location = Find_Entity_Cell(WorkBook_Info, Error_Message_Entity)
        
        Err0 = Split(Error_Message_Components(4), ":")
        Err1 = Split(Err0(0), "/")
        Err2 = Split(Err1(2), " ")
    End If

    Set BreakUp_Error_Msg = ret
End Function

Private Function BreakUp_Error_Msg_DEV(ByRef WorkBook_Info As Collection, ByVal Error_Message As String) As Collection
    'This assumes E3 message format has not changed
    Dim ret As New Collection
    
    Const E3_Clue_String As String = "Received FDC excursion - "
    Const Tool_Alarm_String As String = "TEM: Publishing to FabSnoop the Alarm: "
    Const General_Error_String As String = "rror" 'NO 'e' in 'error' included to avoid lower/upper case
    
    Dim E3_Error_Type As Long
    E3_Error_Type = 0
    
    Dim Error_Message_Entity As String
        
    E3_Error_Type = InStr(1, Error_Message, E3_Clue_String)
    If E3_Error_Type <> 0 Then
        i0 = Split(Error_Message, E3_Clue_String)
        i1 = Split(i0(1), " ")
        i2 = 0
        i3 = 0
        i4 = 0
        i5 = 0
        i6 = 0
        i7 = 0
        i8 = 0
    End If

    Set BreakUp_Error_Msg_DEV = ret
End Function

Function Clean_Up_Input(Input_String As String) As String
    Dim ret As String
    ret = ""
    If Input_String <> "" Then
        Split_Input_Parts = Split(Input_String, vbCrLf)
        Split_Input_Parts = Split_Input_Parts(0)
        Split_Input_Parts = Split(Split_Input_Parts, vbCr)
        Split_Input_Parts = Split_Input_Parts(0)
        Split_Input_Parts = Split(Split_Input_Parts, vbLf)
        Split_Input_Parts = Split_Input_Parts(0)
        Split_Input_Parts = Split(Split_Input_Parts, vbTab)
        Split_Input_Parts = Split_Input_Parts(0)
        Split_Input_Parts = Split(Split_Input_Parts, Chr(160))
        Split_Input_Parts = Split_Input_Parts(0)
        Split_Input_Parts = Split(Split_Input_Parts, " ")
        ret = Split_Input_Parts(0)
    End If
    
    Clean_Up_Input = ret
End Function

Private Function Find_Entity_Cell(ByRef WorkBook_Info As Collection, ByVal Entity As String) As String
    Dim ret As String
    ret = "SKIP"
    
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = "Tool Status"
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = 597
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = 7

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
    Find_Entity_Cell = ret
End Function

Function Already_Compared_Check(ByRef Compared_Collection As Collection, ByVal Entity As String) As Boolean
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

