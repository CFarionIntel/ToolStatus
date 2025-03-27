Attribute VB_Name = "WOPR_Update"
Sub Testing_123()
Attribute Testing_123.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Last_Row As Long
    Last_Row = ActiveSheet.Range("A65535").End(xlUp).Row
    
    For Row = 1 To Last_Row
        Input_String = ActiveSheet.Cells(Row, 1)
        
        comp = Split(Input_String, " ")
        num = UBound(comp)
        
        
        cqt_time = Split(comp(num - 1), "hr")(0)
        start_op = Split(comp(num), "-")(0)
        end_op = Split(comp(num), "-")(1)
        
    Next Row
    
    
    
End Sub
