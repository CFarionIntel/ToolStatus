Attribute VB_Name = "WorkBook_Setups"
Public Function Tool_Status_WorkBook_Setup() As Collection
    a2l "Adding Tool Status Workbook Parameters..."
    ' ToolSts_DashBoard_SheetName
    ' DashBoard_Entity_Column
    ' DashBoard_Last_Row
    ' DashBoard_CEID_Column
    ' DashBoard_Comments_Column
    ' DashBoard_Module_Column
    ' DashBoard_First_WOPR_Column
    ' TEMP_CELL_HOLDER 
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
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
    
    Dim DashBoard_First_WOPR_Column As Long 'Today's comments column
    DashBoard_First_WOPR_Column = Global_Functions.Find_Col("WOPR ID", ToolSts_DashBoard_SheetName)

    STRING_Range_Entity_Col = Global_Functions.Column2Letter(DashBoard_Entity_Column) & "1:" & Global_Functions.Column2Letter(DashBoard_Entity_Column) & DashBoard_Last_Row
        
    'VITAL: This MUST be in alphabetical order or else the algorithms do not work
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
        .Add DashBoard_Entity_Column
        .Add DashBoard_Last_Row
        .Add DashBoard_CEID_Column
        .Add DashBoard_Comments_Column
        .Add DashBoard_Module_Column
        .Add DashBoard_First_WOPR_Column
        .Add TEMP_CELL_HOLDER
    End With
    
    Set Tool_Status_WorkBook_Setup = ret_Component_Names
    a2l "Added!"
End Function

Public Function Abort_History_WorkBook_Setup() As Collection
    ' Abort_FullHistory_SheetName
    ' Abort_History_Entity_Column
    ' Abort_History_Lot_Column
    ' Abort_History_Operation_Column
    ' Abort_History_Slot_Column
    ' Abort_History_WAF3_Column
    ' Abort_History_ChamberPath_Column
    ' Abort_History_Recipe_Column
    ' Abort_History_Start_Column
    ' Abort_History_End_Column
    ' Abort_History_ProcessTime_Column
    ' TEMP_CELL_HOLDER
    
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
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
    
    Set Abort_History_WorkBook_Setup = ret_Component_Names
End Function

Public Function Updated_WOPRs_During_Shift_WorkBook_Setup() As Collection
    ' Passdown_WOPR_SheetName
    ' Passdown_Last_Row
    ' Passdown_Entity_Column
    ' Passdown_CEID_Column
    ' Passdown_State_Column
    ' Passdown_WOPR_Column
    ' Passdown_Status_Column
    ' Passdown_Prio_Column
    ' Passdown_Last_Updated_Column
    ' Passdown_Description_Column
    ' TEMP_CELL_HOLDER
        
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const Passdown_WOPR_SheetName As String = "Passdown"
    
    'Prep Output Sheet
    On Error Resume Next
        ActiveWorkbook.Worksheets(Passdown_WOPR_SheetName).ShowAllData
    
    Dim Passdown_Last_Row As Long
    Passdown_Last_Row = ActiveWorkbook.Worksheets(Passdown_WOPR_SheetName).Range("A65535").End(xlUp).Row
    
    If Passdown_Last_Row = 1 Then
        Passdown_Last_Row = 2
    End If
    
    ClearContentsRange = "A2:" & Global_Functions.Column2Letter(8) & Passdown_Last_Row
    ActiveWorkbook.Worksheets(Passdown_WOPR_SheetName).Range(ClearContentsRange).ClearContents
    
    Dim Passdown_Entity_Column As Long
    Passdown_Entity_Column = Global_Functions.Find_Col("ENTITY", Passdown_WOPR_SheetName)
    
    Dim Passdown_CEID_Column As Long
    Passdown_CEID_Column = Global_Functions.Find_Col("CEID", Passdown_WOPR_SheetName)
    
    Dim Passdown_State_Column As Long
    Passdown_State_Column = Global_Functions.Find_Col("STATE", Passdown_WOPR_SheetName)
    
    Dim Passdown_WOPR_Column As Long
    Passdown_WOPR_Column = Global_Functions.Find_Col("WOPR", Passdown_WOPR_SheetName)
    
    Dim Passdown_Status_Column As Long
    Passdown_Status_Column = Global_Functions.Find_Col("STATUS", Passdown_WOPR_SheetName)
    
    Dim Passdown_Prio_Column As Long
    Passdown_Prio_Column = Global_Functions.Find_Col("PRIO", Passdown_WOPR_SheetName)
    
    Dim Passdown_Last_Updated_Column As Long
    Passdown_Last_Updated_Column = Global_Functions.Find_Col("DATE", Passdown_WOPR_SheetName)
    
    Dim Passdown_Description_Column As Long
    Passdown_Description_Column = Global_Functions.Find_Col("DESC", Passdown_WOPR_SheetName)
    
    Passdown_Last_Row = 2

    With ret_Component_Names
        .Add Passdown_WOPR_SheetName
        .Add Passdown_Last_Row
        .Add Passdown_Entity_Column
        .Add Passdown_CEID_Column
        .Add Passdown_State_Column
        .Add Passdown_WOPR_Column
        .Add Passdown_Status_Column
        .Add Passdown_Prio_Column
        .Add Passdown_Last_Updated_Column
        .Add Passdown_Description_Column
        .Add TEMP_CELL_HOLDER
    End With
    
    Set Updated_WOPRs_During_Shift_WorkBook_Setup = ret_Component_Names
End Function

Public Function Abort_Input_WorkBook_Setup() As Collection
    ' Abort_Setup_Sheetname
    ' Abort_Input_Column
    ' Lot_Input_Row
    ' CurrentOp_Input_Row
    ' SafeMergeOp_Input_Row
    ' Tool_Input_Row
    ' QEF_Input_Row
    ' ErrorMessage_Input_Row
    ' DaysBack_Input_Row
    ' Abort_Output_Column
    ' Lot_Output_Row
    ' Operation_Output_Row
    ' Entity_Output_Row
    ' MMO_Output_Row
    ' RMI_Output_Row
    ' Partial_Output_Row
    ' Error_Output_Row
    ' Output_Col
    ' Abort_WOPR_Title_Row
    ' Teams_Message_Row
    ' Chamber_WOPR_Title_Row
    ' AMF4_Row
    ' TEMP_CELL_HOLDER
    
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const Abort_Setup_Sheetname As String = "Abort Setup"
    Const Abort_Input_Column As Long = 2
    Const Lot_Input_Row As Long = 1
    Const CurrentOp_Input_Row As Long = 2
    Const SafeMergeOp_Input_Row As Long = 3
    Const Tool_Input_Row As Long = 4
    Const QEF_Input_Row As Long = 5
    Const ErrorMessage_Input_Row As Long = 6
    Const DaysBack_Input_Row As Long = 7
    Const Abort_Output_Column As Long = 4
    Const Lot_Output_Row As Long = 11
    Const Operation_Output_Row As Long = 12
    Const Entity_Output_Row As Long = 13
    Const MMO_Output_Row As Long = 15
    Const RMI_Output_Row As Long = 16
    Const Partial_Output_Row As Long = 17
    Const Error_Output_Row As Long = 18
    Const Output_Col As Long = 1
    Const Abort_WOPR_Title_Row As Long = 47
    Const Teams_Message_Row As Long = 48
    Const Chamber_WOPR_Title_Row As Long = 49
    Const AMF4_Row As Long = 50
    
    With ret_Component_Names
        .Add Abort_Setup_Sheetname
        .Add Abort_Input_Column
        .Add Lot_Input_Row
        .Add CurrentOp_Input_Row
        .Add SafeMergeOp_Input_Row
        .Add Tool_Input_Row
        .Add QEF_Input_Row
        .Add ErrorMessage_Input_Row
        .Add DaysBack_Input_Row
        .Add Abort_Output_Column
        .Add Lot_Output_Row
        .Add Operation_Output_Row
        .Add Entity_Output_Row
        .Add MMO_Output_Row
        .Add RMI_Output_Row
        .Add Partial_Output_Row
        .Add Error_Output_Row
        .Add Output_Col
        .Add Abort_WOPR_Title_Row
        .Add Teams_Message_Row
        .Add Chamber_WOPR_Title_Row
        .Add AMF4_Row
        .Add TEMP_CELL_HOLDER
    End With
    
    Set Abort_Input_WorkBook_Setup = ret_Component_Names
End Function

Public Function New_Abort_Input_WorkBook_Setup() As Collection
    ' Abort_Setup_Sheetname
    ' Data_Type_Column
    ' Data_Column
    ' Error_Message_Input_Row
    ' Manual_Override_Lot
    ' Manual_Override_Operation
    ' Manual_Override_Entity
    ' Manual_Override_DaysBack
    ' Abort_WOPR_Title_Row
    ' Teams_Message_Row
    ' Chamber_WOPR_Title_Row
    ' AMF4_Row
    ' Search_Chambers_Row
    ' Search_Hours_Back_Row
    ' TEMP_CELL_HOLDER
    
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const Abort_Setup_Sheetname As String = "New_Abort_Input"
    
    Dim Data_Type_Column As Long
    Data_Type_Column = Global_Functions.Find_Col("Column Type", Abort_Setup_Sheetname)
    
    Dim Data_Column As Long
    Data_Column = Global_Functions.Find_Col("Column Data", Abort_Setup_Sheetname)
    
    Const Error_Message_Input_Row As Long = 2
    Const Manual_Override_Lot As Long = 5
    Const Manual_Override_Operation As Long = 6
    Const Manual_Override_Entity As Long = 7
    Const Manual_Override_DaysBack As Long = 8
    Const Abort_WOPR_Title_Row As Long = 31
    Const Teams_Message_Row As Long = 32
    Const Chamber_WOPR_Title_Row As Long = 33
    Const AMF4_Row As Long = 34
    Const Search_Chambers_Row As Long = 15
    Const Search_Hours_Back_Row As Long = 16
    
    With ret_Component_Names
        .Add Abort_Setup_Sheetname
        .Add Data_Type_Column
        .Add Data_Column
        .Add Error_Message_Input_Row
        .Add Manual_Override_Lot
        .Add Manual_Override_Operation
        .Add Manual_Override_Entity
        .Add Manual_Override_DaysBack
        .Add Abort_WOPR_Title_Row
        .Add Teams_Message_Row
        .Add Chamber_WOPR_Title_Row
        .Add AMF4_Row
        .Add Search_Chambers_Row
        .Add Search_Hours_Back_Row
        .Add TEMP_CELL_HOLDER
    End With
    
    Set New_Abort_Input_WorkBook_Setup = ret_Component_Names
End Function

Public Function Tool_Status_History_WorkBook_Setup() As Collection
    a2l "Adding Tool Status History Workbook Parameters..."
    ' ToolSts_FullHistory_SheetName
    ' History_Last_Column
    
    Dim ret_Component_Names As New Collection
    
    Const ToolSts_FullHistory_SheetName As String = "ToolStsHistory"
    
    Dim History_Last_Column As Long
    History_Last_Column = (ActiveWorkbook.Worksheets(ToolSts_FullHistory_SheetName).Range("CLF1").End(xlToLeft).Column + 1)
    
    With ret_Component_Names
        .Add ToolSts_FullHistory_SheetName
        .Add History_Last_Column
    End With
    
    Set Tool_Status_History_WorkBook_Setup = ret_Component_Names
    a2l "Added!"
End Function

Public Function Change_Report_WorkBook_Setup() As Collection
    a2l "Adding Change Report Workbook Parameters..."
    ' Change_Report_SheetName
    ' Last_Change_Report_Row
    ' Change_Report_Time_Col
    ' Change_Report_UTP_Col
    ' Change_Report_Down_Col

    Dim ret_Component_Names As New Collection
    
    Const Change_Report_SheetName As String = "Change Report"
    
    Dim Last_Change_Report_Row As Long
    Last_Change_Report_Row = ActiveWorkbook.Sheets(Change_Report_SheetName).Range("A1812").End(xlUp).Row + 1
    
    Const Change_Report_Time_Col As Long = 1
    Const Change_Report_UTP_Col As Long = 2
    Const Change_Report_Down_Col As Long = 3
    
    With ret_Component_Names
        .Add Change_Report_SheetName
        .Add Last_Change_Report_Row
        .Add Change_Report_Time_Col
        .Add Change_Report_UTP_Col
        .Add Change_Report_Down_Col
    End With
    
    Set Change_Report_WorkBook_Setup = ret_Component_Names
    a2l "Added!"
End Function

Public Function CLUI_WorkBook_Setup() As Collection
    ' CLUI_SheetName
    ' CLUI_Date_Column
    ' CLUI_Last_Row
    ' CLUI_Event_Column
    ' CLUI_State_Column
    ' CLUI_Comments_Column
    ' CLUI_User_Column
    ' TEMP_CELL_HOLDER
    Dim ret_Component_Names As New Collection
    
    Dim TEMP_CELL_HOLDER As String 'This will be temporary storage for a cell location in other functions
    TEMP_CELL_HOLDER = "A1"
    
    Const CLUI_SheetName As String = "CLUI"
    Const CLUI_Starting_Row As Long = 5
    Const CLUI_Starting_Col As Long = 3
    
    'Prep Output Sheet
    On Error Resume Next
        ActiveWorkbook.Worksheets(CLUI_SheetName).ShowAllData
    
    Const CLUI_Entry_Column As Long = 2
    
    Dim CLUI_Date_Column As Long
    CLUI_Date_Column = Global_Functions.Find_Col("Date Time", CLUI_SheetName, CLUI_Starting_Row, CLUI_Starting_Col)
    
    Dim CLUI_Event_Column As Long 'Today's comments column
    CLUI_Event_Column = Global_Functions.Find_Col("Event", CLUI_SheetName, CLUI_Starting_Row, CLUI_Starting_Col)
    
    Dim CLUI_Last_Row As Long
    CLUI_Last_Row = ActiveWorkbook.Worksheets(CLUI_SheetName).Range(Global_Functions.Column2Letter(CLUI_Date_Column) & "65535").End(xlUp).Row
    
    Dim CLUI_State_Column As Long
    CLUI_State_Column = Global_Functions.Find_Col("State", CLUI_SheetName, CLUI_Starting_Row, CLUI_Starting_Col)
    
    Dim CLUI_Comments_Column As Long
    CLUI_Comments_Column = Global_Functions.Find_Col("Comments", CLUI_SheetName, CLUI_Starting_Row, CLUI_Starting_Col)
    
    Dim CLUI_User_Column As Long 'Today's comments column
    CLUI_User_Column = Global_Functions.Find_Col("User", CLUI_SheetName, CLUI_Starting_Row, CLUI_Starting_Col)
    
    With ret_Component_Names
        .Add CLUI_SheetName
        .Add CLUI_Date_Column
        .Add CLUI_Last_Row
        .Add CLUI_Event_Column
        .Add CLUI_State_Column
        .Add CLUI_Comments_Column
        .Add CLUI_User_Column
        .Add TEMP_CELL_HOLDER
    End With
    
    Set CLUI_WorkBook_Setup = ret_Component_Names
End Function
