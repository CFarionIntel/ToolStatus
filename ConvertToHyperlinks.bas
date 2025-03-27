Attribute VB_Name = "ConvertToHyperlinks"
Sub ConvertRangeToHyperlinks(rng As Range)
    Dim cell As Range
    
    Const WOPR_Link_Address As String = "https://rf3-apps-fuzion.rf3prod.mfg.intel.com/EditWorkOrderPage.aspx?WorkOrderId="
    
    ' Check if the range is set
    If rng Is Nothing Then
        MsgBox "No range provided. Exiting macro.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value <> "" And Not cell.Hyperlinks.Count > 0 Then ' Check if the cell is not empty and does not already contain a hyperlink
            ' Create a hyperlink in the cell that links to itself
            cell.Parent.Hyperlinks.Add Anchor:=cell, Address:=WOPR_Link_Address & cell.Text, TextToDisplay:=cell.Text
        End If
    Next cell
End Sub

Sub CallConvertRangeToHyperlinks()
Attribute CallConvertRangeToHyperlinks.VB_ProcData.VB_Invoke_Func = "L\n14"
    Dim targetRange As Range
    
    ' Set the target range to the desired cells
    Set targetRange = Selection
    
    ' Call the subroutine to convert the range to hyperlinks
    ConvertRangeToHyperlinks targetRange
End Sub
