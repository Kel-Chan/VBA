Attribute VB_Name = "checkboxes"

Sub addcheckboxes()

    Dim cell As Range
    
    For Each cell In Selection
        
        Set cb = ActiveSheet.checkboxes.Add(cell.Left, cell.Top, 60, 20)
        cb.Caption = "Apple"
        cb.LinkedCell = cell.Address
        cell.NumberFormat = ";;;"
        
        
    Next cell
End Sub

Sub delectCheckboxes()
    For Each cb In ActiveSheet.checkboxes
        cb.Delete
    Next cb
End Sub
