Attribute VB_Name = "Main_CEID_Check"
Sub Update_Check_CEID()
    Debug.Print ("Start CEID Check > " & Format(Time, "hh:mm.ss"))
    Application.StatusBar = True
    
    Dim Current_CEID As String
    Dim New_CEID As String
    Dim Difference As Boolean

    'Collect the Sheet Names and columns which comprise the Dashboard
    Dim WorkBook_Component_Names As Collection
    Set WorkBook_Component_Names = Tool_Status_WorkBook_Setup
    
    'Select how you want to collect the new Tool Status Update
    Dim ToolSts_Input_Selection As Long
    ToolSts_Input_Selection = ActiveWorkbook.Sheets("Settings").Cells(1, 1)
    
    'Create a collection of the latest Tool Status from the Input
    Dim Entries As Collection
    Set Entries = Create_Latest_ToolSts_Collection(ToolSts_Input_Selection)
    
    If Entries.Count <> 0 Then
        'Processing for each entry in Query. This will output a Change Report
        Dim Change_Report As New Collection
        Set Change_Report = Process_for_Each_Entry(WorkBook_Component_Names, Entries)
        
        'Take Change Report and output to Change Report Sheets
        Record_Changes WorkBook_Component_Names, Change_Report
    End If
    
    Application.StatusBar = False
    Debug.Print ("End CEID Check > " & Format(Time, "hh:mm.ss"))
End Sub

Private Function Process_for_Each_Entry(ByRef WorkBook_Info As Collection, ByRef Entries As Collection) As Collection
    Dim Change_Report As New Collection
    Dim Entity_Cell_Location As String
    
    For Entry = 1 To Entries.Count
        Entity_Cell_Location = Find_Entity_Cell(WorkBook_Info, Entries(Entry).Name)
        If Entity_Cell_Location <> "SKIP" Then
            Add_Change_to_Report WorkBook_Info, Entries(Entry), Change_Report 'Update Change Report (Currently in main function)
        End If
    Next Entry
    
    Set Process_for_Each_Entry = Change_Report
End Function

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

Private Sub Add_Change_to_Report(ByRef WorkBook_Info As Collection, ByVal Entity_Entry As Entity, ByRef Change_Report As Collection)
    Const DEV As Boolean = True
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim CEID_Col As Long
    CEID_Col = Find_Col("CEID", ToolSts_DashBoard_SheetName)
    
    Dim Message As String
    Message = ""
    
    Dim Current_CEID As String
    Current_CEID = ActiveWorkbook.Worksheets(ToolSts_DashBoard_SheetName).Cells(Range(WorkBook_Info(WorkBook_Info.Count)).Row, CEID_Col)
    
    Dim New_CEID As String
    New_CEID = Entity_Entry.CEID
    
    If Current_CEID <> New_CEID Then
        Message = Entity_Entry.Name & ":" & Current_CEID & ":" & New_CEID
        Change_Report.Add Message
    End If
End Sub

Private Function Create_Latest_ToolSts_Collection(Input_Option As Long)
    Dim Entries As New Collection
    
    If Input_Option = 1 Then
        'Create Collection from Sheet Values
    ElseIf Input_Option = 2 Then
        Dim STRING_SQL_OUTPUT_FILE As String
        STRING_SQL_OUTPUT_FILE = "C:\Users\cfarion\AppData\Local\Temp\SQLPathFinder_Temp\out_SQL_Tool_Status.tab"
        
        Dim Text As String
        Text = Global_Functions.Read_SQL_Output_File(STRING_SQL_OUTPUT_FILE)
        
        Populate_Entry_Collection Entries, Text
    ElseIf Input_Option = 3 Then
        SQL_Script_Through_Excel Entries
    Else
        'Nothing Selected or Error
    End If

    Set Create_Latest_ToolSts_Collection = Entries
End Function

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

Private Sub Record_Changes(ByRef WorkBook_Info As Collection, ByRef Changes As Collection)
    Dim ToolSts_DashBoard_SheetName As String
    ToolSts_DashBoard_SheetName = WorkBook_Info(1)
    
    Dim DashBoard_Entity_Column As Long
    DashBoard_Entity_Column = WorkBook_Info(3)
    
    Dim DashBoard_Last_Row As Long
    DashBoard_Last_Row = WorkBook_Info(4)
    
    Const CEID_Check_SheetName As String = "CEID Check"
    
    Const Max_Lines_for_MsgBox As Long = 12
    
    Dim Num_of_Changes As Long
    Num_of_Changes = Changes.Count
    
    'Dim Change_String As String
    Change_String = ""
    
    Time_Format = Format(Date, "mm/dd/yyyy") & " - " & Format(Time, "hh:mm.ss")
    
    'Dim Last_Change_Report_Row As Long
    Last_Change_Report_Row = ActiveWorkbook.Sheets(CEID_Check_SheetName).Range("A1812").End(xlUp).Row + 1
    
    If Num_of_Changes <> 0 Then
        Change_String = "CEID Updates Available!!!"
        For Current_Change = 1 To Num_of_Changes
            Components = Split(Changes(Current_Change), ":")
            
            With Sheets(CEID_Check_SheetName)
                .Cells(Last_Change_Report_Row + Current_Change - 1, 1) = Time_Format
                .Cells(Last_Change_Report_Row + Current_Change - 1, 2) = Components(0)
                .Cells(Last_Change_Report_Row + Current_Change - 1, 3) = Components(1)
                .Cells(Last_Change_Report_Row + Current_Change - 1, 4) = Components(2)
            End With
            
            
        Next Current_Change
    Else
        Change_String = "No CEID Changes"
    End If
    
    msg = MsgBox(Change_String, 0, Time_Format, 0, 0)
End Sub

