Attribute VB_Name = "New_Main_Abort"
Sub New_Abort_Processing()
    Debug.Print ("Start New Abort > " & Format(Time, "hh:mm.ss"))
    'Collect the Sheet Names and columns which comprise the Dashboard
    
    Dim WorkBook_Input_Parameters As Collection
    Set WorkBook_Input_Parameters = New_Abort_Input_WorkBook_Setup
    
    Dim Entry As Abort
    Dim Entries As New Collection
    Dim Manual_Override_Needed As Boolean
    Manual_Override_Needed = False
    
    Dim errStr As String
    'errStr = "Received FDC excursion - LotId:D4374260, SubEntity:GTO413_PC4, WaferId:BHFKB193SJF5, ErrorMsg:Process_Step05/RFVppUp/Stdev Result=25.095 State=CRITICAL Region=UPPER_CRITICAL Upper_Critical=21.604 Upper_Warning=16.729 Target=6.979 Lower_Warning=-2.771 Lower_Critical=-7.646 Collection=P1278_GTOcn_CE_BOHR200_BVx_PROD_NEW_BKM ContextGroup=DUMMY"
    errStr = ActiveWorkbook.Sheets(WorkBook_Input_Parameters(1)).Cells(WorkBook_Input_Parameters(4), WorkBook_Input_Parameters(3)).Value
    
    Application.StatusBar = "Parsing error message..."
    Dim Init_Lot_Tool_and_Error As Collection
    Set Init_Lot_Tool_and_Error = Parse_Error_Msg(errStr)
    
    'Set Init_Lot_Tool_and_Error = Unit_Testing_00(Init_Lot_Tool_and_Error)
    Application.StatusBar = "Collecting lots from initial search..."
    Dim Lots_Recently_Processed As New Collection
    SQL_Script_for_Initial_Search Lots_Recently_Processed, Init_Lot_Tool_and_Error
    
    'Set Lots_Recently_Processed = Unit_Testing_01(Lots_Recently_Processed)
    Application.StatusBar = "Determining the correct information for the SQL script..."
    
    If Lots_Recently_Processed(1).DateStart <> "No Start Date" Then
        Set Entry = Determine_Correct_Entry(Lots_Recently_Processed)
        If Entry.Lot = "Multiple Prod lots. Use Manual Override" Then
            'ActiveWorkbook.Sheets(WorkBook_Input_Parameters(1)).Cells(2, 3) = "Manual Override Needed"
            Application.StatusBar = False
            Debug.Print ("End Main > " & Format(Time, "hh:mm.ss"))
            MsgBox "Multiple Lots on record. Please try Manual Override."
            Exit Sub
        End If
        
        WorkBook_Input_Parameters.Add Entry
        
        SQL_Script_Through_Excel Entries, WorkBook_Input_Parameters
    Else
        MsgBox "No Lots on record. Please try Manual Override."
        Exit Sub
    End If
    
    If Entries.Count <> 0 Then
        Dim WorkBook_Output_Parameters As Collection
        Set WorkBook_Output_Parameters = Abort_History_WorkBook_Setup
        
        'What is typically displayed from a SQL script
        Record_Query Entries, WorkBook_Output_Parameters
        
        'Dispo each wafer
        Dispo_Wafers Entries
        Display_Abort_Statistics Entries, WorkBook_Output_Parameters
        
        Dim WOPR_Components As New Collection
        With WOPR_Components
            .Add errStr 'Full Error String
            .Add Init_Lot_Tool_and_Error(2) 'Full SubEntity
            .Add Init_Lot_Tool_and_Error(3) 'Condensed Error Message
            .Add Entry.Lot 'PROD lot
            .Add Entries(1).Get_AMF4_Recipe
        End With
        
        E3_Message_Output WorkBook_Input_Parameters, WOPR_Components
        
        Get_Chambers_for_E3_Search WorkBook_Input_Parameters, Entries
        'Get Hours_Back_for_E3_Search WorkBook_Input_Parameters, Entries
    End If
    
    Application.StatusBar = False
    Debug.Print ("End Main > " & Format(Time, "hh:mm.ss"))
End Sub

Sub Get_Chambers_for_E3_Search(ByRef WorkBook_Info As Collection, ByRef Entries As Collection)
    Dim ret_Chamber_String As String
    Dim ret_Time_String As String
    Dim in_chamber_String As Boolean
    in_chamber_String = False
    Dim Hours_to_SearchBack As Double
    Hours_to_SearchBack = 12
    
    ret_Chamber_String = ""
    
    
    If Entries(1).DateStart <> "No Start Date" Then
        CurrentTime_Difference = (Date + Time) - (DateValue(Entries(1).DateStart) + TimeValue(Entries(1).DateStart))
        Hours_to_SearchBack = Round(24 * CurrentTime_Difference, 0) + 1
    End If
    
    ActiveWorkbook.Sheets(WorkBook_Info(1)).Cells(WorkBook_Info(14), WorkBook_Info(3)) = Hours_to_SearchBack & vbCrLf & Entries(1).DateStart
    
    If Entries.Count <> 0 Then
        For i = 1 To Entries.Count
            in_chamber_String = False
            i0 = Split(ret_Chamber_String, ";")
            strt_pt = InStr(1, Entries(i).ChamberPath, "PM")
            If strt_pt <> 0 Then
                Chamber_Num = Mid(Entries(i).ChamberPath, strt_pt, 3)
                For j = 0 To UBound(i0)
                    If i0(j) = Chamber_Num Then
                        in_chamber_String = True
                    End If
                Next j
            End If
            If Not in_chamber_String And strt_pt <> 0 Then
                ret_Chamber_String = ret_Chamber_String & Chamber_Num & ";"
            End If
        Next i
    End If
    ActiveWorkbook.Sheets(WorkBook_Info(1)).Cells(WorkBook_Info(13), WorkBook_Info(3)) = ret_Chamber_String
End Sub

Private Sub E3_Message_Output(ByRef WorkBook_Info As Collection, ByRef Error_Info As Collection)
    Dim ret As String
    
    Dim Ent_WOPR As New WOPR
    Dim Module As String
    
    Dim Dashboard_WorkBook_Parameters As Collection
    Set Dashboard_WorkBook_Parameters = Tool_Status_WorkBook_Setup
    
    Full_Error = Error_Info(1)
    SubEntity = Error_Info(1)
    Condensed_Error_Message = Error_Info(3)
    Prod_Lot = Error_Info(4)
    
    Dim Abort_Setup_Sheetname As String
    Abort_Setup_Sheetname = WorkBook_Info(1)
    
    Const Title_Format As String = "[..::..] POR Lot Abort Recovery - HB? Lot ..::.. - ..::.."
    
    ts_ = Split(Title_Format, "..::..")
    
    ret = Find_Entity_Cell(Dashboard_WorkBook_Parameters, Error_Info(2))
    
    If ret <> "SKIP" Then
        Module = Sheets(Dashboard_WorkBook_Parameters(1)).Cells(Range(ret).Row, 1)
        Bay = Sheets(Dashboard_WorkBook_Parameters(1)).Cells(Range(ret).Row, 6)
    End If
    
    AMF4_Route = ""
    AMF4_Entity = Error_Info(2)
    AMF4_Event = "4P" & Right(AMF4_Entity, 1) & "_ETCH_TEST"
    AMF4_Chamber = Right(SubEntity, 3)
    
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
    
    'AMF4_Entity = Error_Info(2)
    'AMF4_Event = "4P" & Right(AMF4_Entity, 1) & "_ETCH_TEST"
    AMF4_Recipe = Error_Info(5)
    OutputMsg = ts_(0) & Module & ts_(1) & Prod_Lot & ts_(2) & Condensed_Error_Message & Chr(10) & Chr(10) & Full_Error
    OutputMsg1 = "[" & Module & "] " & Condensed_Error_Message & " Recovery" & Chr(10) & Chr(10) & Full_Error
    AMF4 = "Route: " & AMF4_Route & Chr(10) & "Entity: " & AMF4_Entity & Chr(10) & "Event: " & AMF4_Event & Chr(10) & "Recipe: " & AMF4_Recipe & Chr(10) & "Chamber: " & AMF4_Chamber & Chr(10) & "Lot: " & Chr(10) & "Slots: Any 3"
    Worksheets(Abort_Setup_Sheetname).Cells(31, 2) = OutputMsg
    Worksheets(Abort_Setup_Sheetname).Cells(32, 2) = "Hi team, here is WO# 3811495 for the non-HB and non-CQT abort on " & AMF4_Entity & " (" & Bay & "). Thank you!"
    Worksheets(Abort_Setup_Sheetname).Cells(33, 2) = OutputMsg1
    Worksheets(Abort_Setup_Sheetname).Cells(34, 2) = AMF4
    
End Sub

Private Function Determine_Correct_Entry(ByRef Search_Entries As Collection)
    Dim Entry As Abort
    Set Entry = New Abort
    
    Dim temp_Collection As New Collection
    
    Dim Entry_Num As Long
    Entry_Num = Search_Entries.Count
    
    Dim Multiple_Prod_Lots As Boolean
    Multiple_Prod_Lots = False
    
    Dim Lot_String As String
    Lot_String = ""
    
    If Entry_Num <> 0 Then
        For i = 1 To Entry_Num
            TW = Mid(Search_Entries(i).Lot, 5, 1)
            DCS = InStr(Search_Entries(i).Lot, "_DCS")
            If TW <> "T" And DCS = 0 Then
                temp_Collection.Add Search_Entries(i)
            End If
        Next i
        
        If temp_Collection.Count > 1 Then
            For i = 1 To temp_Collection.Count - 1
                If temp_Collection(i).Lot <> temp_Collection(i + 1).Lot Then
                    Multiple_Prod_Lots = True
                End If
            Next i
            
            If Multiple_Prod_Lots = True Then
                Entry.Lot = "Multiple Prod lots. Use Manual Override"
                Set Determine_Correct_Entry = Entry
                Exit Function
            End If
            Entry.Lot = temp_Collection(1).Lot
            Entry.Entity = temp_Collection(1).Entity
            Entry.Operation = temp_Collection(1).Operation
        Else
            Entry.Lot = temp_Collection(1).Lot
            Entry.Entity = temp_Collection(1).Entity
            Entry.Operation = temp_Collection(1).Operation
        End If
    End If
    
    Set Determine_Correct_Entry = Entry
End Function

Private Sub SQL_Script_for_Initial_Search(ByRef coll As Collection, ByRef Search_Info As Collection)
    On Error GoTo ErrHandler
    
    Dim DataSource As String
    DataSource = "D1D_PROD_XEUS"

    Dim sQuery As String
    
    Dim Input_Query As String
    Input_Query = ActiveWorkbook.Sheets("SQL_INPUT").Cells(15, 2).Value
    
    sQuery = Build_First_Search_Query(Search_Info, Input_Query)
    
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
        Entry.Entity = UCase(Search_Info(2))
        Entry.Lot = Search_Info(1)
        Entry.Operation = 0
        Entry.Slot = 0
        Entry.Waf3 = 0
        Entry.ChamberPath = "Has not entered Entity in the past 20 minutes"
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
    
    Const Abort_Input_SheetName As String = "New_Abort_Input"
    Const Abort_FullHistory_SheetName As String = "AbortHistory"
    
    Dim Data_Type_Column As Long
    Data_Type_Column = Global_Functions.Find_Col("Column Type", Abort_Input_SheetName)
    
    Dim Data_Column As Long
    Data_Column = Global_Functions.Find_Col("Column Data", Abort_Input_SheetName)
    
    ''' These will be constants for now. Need to dynamically find later
    Dim Input_Message_Row As Long
    Input_Message_Row = 2
    
    Dim Abort_WOPR_Title_Row As Long
    Abort_WOPR_Title_Row = 31
    
    Dim Teams_Message_Row As Long
    Teams_Message_Row = 32
    
    Dim Chamber_WOPR_Title_Row As Long
    Chamber_WOPR_Title_Row = 33
    
    Dim AMF4_Row As Long
    AMF4_Row = 34
    
    Dim Lot_Row As Long
    Lot_Row = 5
    
    Dim Operation_Row As Long
    Operation_Row = 6
    
    Dim Entity_Row As Long
    Entity_Row = 7
            
    With ret_Component_Names
        .Add Abort_Input_SheetName
        .Add Abort_FullHistory_SheetName
        .Add Data_Type_Column
        .Add Data_Column
        .Add Input_Message_Row
        .Add Abort_WOPR_Title_Row
        .Add Teams_Message_Row
        .Add Chamber_WOPR_Title_Row
        .Add AMF4_Row
        .Add Lot_Row
        .Add Operation_Row
        .Add Entity_Row
        .Add TEMP_CELL_HOLDER
    End With
    
    Set WorkBook_Setup = ret_Component_Names
End Function

Private Function Build_First_Search_Query(ByRef Search_Components As Collection, ByVal iQuery As String) As String
    Dim ret_Finished_Query As String
    ret_Finished_Query = ""
    
    Dim Input_Values As New Collection
    
    TW = Mid(Search_Components(1), 5, 1)
    DCS = InStr(Search_Components(1), "_DCS")
    If TW = "T" Or DCS = 0 Then
        Input_Values.Add Clean_Up_Input(UCase(Search_Components(1)))
    Else
        Input_Values.Add ""
    End If
    
    Dim MOM_Entity As String
    MOM_Entity = Split(Search_Components(2), "_")(0)
    
    Dim SubEntity As String
    SubEntity = Split(Search_Components(2), "_")(1)
    
    With Input_Values
        .Add Clean_Up_Input(UCase(MOM_Entity))
        .Add Clean_Up_Input(UCase(SubEntity))
    End With
    
    
    
    Query_Components = Split(iQuery, "..::..")
    
    Dim Number_of_Components As Long
    Number_of_Components = UBound(Query_Components)
    ret_Finished_Query = ret_Finished_Query & Query_Components(0)
    
    For i = 1 To Number_of_Components
        ret_Finished_Query = ret_Finished_Query & Input_Values(i) & Query_Components(i)
    Next i
        
    Build_First_Search_Query = ret_Finished_Query
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

Public Function Parse_Error_Msg(ByVal Error_Message As String) As Collection
    'This assumes E3 message format has not changed
    Dim ret As New Collection
    Dim Error_Type_Strings As New Collection
    
    Const E3_Clue_String As String = "Received FDC excursion - "
    Const Tool_Alarm_String As String = "TEM: Publishing to FabSnoop the Alarm: "
    Const General_Error_String As String = "rror" 'NO 'e' in 'error' included to avoid lower/upper case
    
    With Error_Type_Strings
        .Add E3_Clue_String
        .Add Tool_Alarm_String
        .Add General_Error_String
    End With

    Dim Error_Message_Present As Long
    Error_Message_Present = 0
    Error_Type = 0
    
    Do While Error_Message_Present = 0 And Error_Type < Error_Type_Strings.Count
        Error_Type = Error_Type + 1
        Error_Message_Present = InStr(1, Error_Message, Error_Type_Strings(Error_Type))
    Loop
    
    Select Case Error_Type
        Case 1 'E3 Error
            Set ret = E3_Error_Search(Error_Message, E3_Clue_String)
        Case 2 'Tool Alarm
            Set ret = Tool_Alarm_Search(Error_Message, Tool_Alarm_String)
        Case 3 'Other Alarm
            ret.Add "N/A" ' No LotId
            ret.Add "N/A" ' No SubEntity
            ret.Add Error_Message
        Case Else
            ret.Add "N/A" ' No LotId
            ret.Add "N/A" ' No SubEntity
            ret.Add Error_Message
    End Select

    Set Parse_Error_Msg = ret
End Function

Private Function E3_Error_Search(ByVal Error_Message As String, E3_Clue_String As String) As Collection
    Dim ret As New Collection
    
    i0 = Split(Error_Message, E3_Clue_String)
    i1 = Split(i0(1), " ")
    
    LotId = Split(i1(0), ":")(1)
    LotId = Split(LotId, ",")(0) 'Get rid of the comma
    SubEntity = Split(i1(1), ":")(1)
    SubEntity = Split(SubEntity, ",")(0) 'Get rid of the comma
    'WaferId = Split(i1(2), ":")(1)
    'WaferId = Split(WaferId, ",")(0) 'Get rid of the comma
    ErrorMsg = Split(i1(3), ":")(1)
    'Result = Split(i1(4), "=")(1)
    'State = Split(i1(5), "=")(1)
    'Region = Split(i1(6), "=")(1)
    'Upper_Critical = Split(i1(7), "=")(1)
    'Upper_Warning = Split(i1(8), "=")(1)
    'Target = Split(i1(9), "=")(1)
    'Lower_Warning = Split(i1(10), "=")(1)
    'Lower_Critical = Split(i1(11), "=")(1)
    'Collection = Split(i1(12), "=")(1)
    'ContextGroup = Split(i1(13), "=")(1)
    
    'ErrorMsg Processing
    eMsg = Split(ErrorMsg, "/")
    If eMsg(0) = "CustomEquation" Then
        ErrorMsg = eMsg(2)
    Else
        ErrorMsg = eMsg(1) & "/" & eMsg(2)
    End If
    
    With ret
        .Add LotId
        .Add SubEntity
        '.Add WaferId
        .Add ErrorMsg
        '.Add Result
        '.Add State
        '.Add Region
        '.Add Upper_Critical
        '.Add Upper_Warning
        '.Add Target
        '.Add Lower_Warning
        '.Add Lower_Critical
        '.Add Collection
        '.Add ContextGroup
    End With
    
    Set E3_Error_Search = ret
End Function

Private Function Tool_Alarm_Search(ByVal Error_Message As String, Tool_Alarm_String As String) As Collection
    Dim ret As New Collection
    
    i0 = Split(Error_Message, Tool_Alarm_String)
    i1 = Split(i0(1), ",")
    AlarmID = Split(i1(0), "=")(1)
    AlarmText = Split(i1(1), "=")(1)
    
    Const LotId As String = "N/A"
    Const SubEntity As String = "N/A"
    Dim ErrorMsg As String
    ErrorMsg = Error_Message
    
    With ret
        '.Add AlarmID
        '.Add AlarmText
        .Add LotId
        .Add SubEntity
        .Add ErrorMsg
    End With
    
    Set Tool_Alarm_Search = ret
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

Function Find_Entity_Cell(ByRef WorkBook_Info As Collection, ByVal Entity As String) As String
    ' ToolSts_DashBoard_SheetName
    ' DashBoard_Entity_Column
    ' DashBoard_Last_Row
    ' DashBoard_CEID_Column
    ' DashBoard_Comments_Column
    ' DashBoard_Module_Column
    ' DashBoard_First_WOPR_Column
    ' TEMP_CELL_HOLDER
    
    Dim ret As String
    ret = "SKIP"
    
    Dim ToolSts_DashBoard_SheetName As String
    'ToolSts_DashBoard_SheetName = "Tool Status"
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_Last_Row As Long
    'DashBoard_Last_Row = 597
    DashBoard_Last_Row = WorkBook_Info(3)
    
    Dim DashBoard_Entity_Column As Long
    'DashBoard_Entity_Column = 7
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
        Entry.Entity = UCase(ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(7, 2).Value)
        Entry.Lot = ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(5, 2).Value
        Entry.Operation = ActiveWorkbook.Sheets(Abort_Input_SheetName).Cells(6, 2).Value
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
    
    sza = WorkBook_Info.Count
    
    Dim Input_Values As New Collection
           
    With Input_Values
        .Add WorkBook_Info(sza).Lot
        .Add WorkBook_Info(sza).Entity
        .Add WorkBook_Info(sza).Operation
        .Add 2
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

Function Unit_Testing_00(ByRef coll As Collection) As Collection
    Dim ret As New Collection

    Const Prod_Lot_00 As String = "D45076E0"
    Const Prod_Lot_01 As String = "D4428120"
    Const Prod_Lot_02 As String = "D432Y620"
    Const Prod_Lot_03 As String = "D43277R0"
    Const TW_Lot_00 As String = "D422TGM0"
    
    Const SubEntity As String = "GTO447_PC6"
    
    Const Condensed_Error_Message As String = "RFVppUp/Mean"
    
    
    ret.Add TW_Lot_00
    ret.Add SubEntity
    ret.Add Condensed_Error_Message
    
    Set Unit_Testing_00 = ret
End Function

Function Unit_Testing_01(ByRef coll As Collection) As Collection
    Dim ret As New Collection
    
    Dim Entry As Abort
    
    Const Prod_Lot_00 As String = "D45076E0"
    Const Prod_Lot_01 As String = "D4428120"
    Const Prod_Lot_02 As String = "D432Y620"
    Const Prod_Lot_03 As String = "D43277R0"
    Const TW_Lot_00 As String = "D422TGM0"
    
    Const SubEntity As String = "GTO447_PC6"
    
    Dim Operation_00 As Long
    Operation_00 = Int((999999 - 100000 + 1) * Rnd + 100000)
    
    Dim Operation_01 As Long
    Operation_01 = Int((999999 - 100000 + 1) * Rnd + 100000)
    
    
    Set Entry = New Abort
    Entry.Lot = Prod_Lot_00
    Entry.Operation = Operation_00
    Entry.Entity = SubEntity
    ret.Add Entry
    
    Set Entry = New Abort
    Entry.Lot = TW_Lot_00
    Entry.Operation = Operation_01
    Entry.Entity = SubEntity
    ret.Add Entry
    
    Set Entry = New Abort
    Entry.Lot = Prod_Lot_00
    Entry.Operation = Operation_00
    Entry.Entity = SubEntity
    ret.Add Entry
    
    Set Entry = New Abort
    Entry.Lot = Prod_Lot_00
    Entry.Operation = Operation_00
    Entry.Entity = SubEntity
    ret.Add Entry
    
    Set Unit_Testing_01 = ret
End Function


