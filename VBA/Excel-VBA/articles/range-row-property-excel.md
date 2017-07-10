---
title: Range.Row Property (Excel)
keywords: vbaxl10.chm144188
f1_keywords:
- vbaxl10.chm144188
ms.prod: excel
api_name:
- Excel.Range.Row
ms.assetid: 3c8d7351-4fc6-748b-c2a8-de3dab4b964e
ms.date: 06/08/2017
---


# Range.Row Property (Excel)

Returns the number of the first row of the first area in the range. Read-only  **Long** .


## Syntax

 _expression_ . **Row**

 _expression_ A variable that represents a **Range** object.


## Example

This example sets the row height of every other row on Sheet1 to 4 points.


```vb
For Each rw In Worksheets("Sheet1").Rows 
    If rw.Row Mod 2 = 0 Then 
        rw.RowHeight = 4 
    End If 
Next rw
```

 **Sample code provided by:** Holy Macro! Books,[Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1) |[About the Contributors](range-row-property-excel.md#AboutContributor)

This example uses the  **BeforeDoubleClick** worksheet event to copy a row of data from one worksheet to another. To run this code, the name of the target worksheet must be in column A. When you double click a cell that contains data, this example gets the target worksheet name from column A and copies the entire row of data into the next available row on the target worksheet. This example accesses the active row using the **Target** keyword.




```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'If the double click occurs on the header row or an empty cell, exit the macro.
    If Target.Row = 1 Then Exit Sub
    If Target.Row > ActiveSheet.UsedRange.Rows.Count Then Exit Sub
    If Target.Column > ActiveSheet.UsedRange.Columns.Count Then Exit Sub
    
    'Override the default double-click behavior with this function.
    Cancel = True
    
    'Declare your variables.
    Dim wks As Worksheet, xRow As Long
    
    'If an error occurs, use inline error handling.
    On Error Resume Next
    
    'Set the target worksheet as the worksheet whose name is listed in the first cell of the current row.
    Set wks = Worksheets(CStr(Cells(Target.Row, 1).Value))
    'If there is an error, exit the macro.
    If Err > 0 Then
        Err.Clear
        Exit Sub
    'Otherwise, find the next empty row in the target worksheet and copy the data into that row.
    Else
        xRow = wks.Cells(wks.Rows.Count, 1).End(xlUp).Row + 1
        wks.Range(wks.Cells(xRow, 1), wks.Cells(xRow, 7)).Value = _
        Range(Cells(Target.Row, 1), Cells(Target.Row, 7)).Value
    End If
End Sub
```

 **Sample code provided by:** Dennis Wallentin,[VSTO &; .NET &; Excel](http://xldennis.wordpress.com/) |[About the Contributors](range-row-property-excel.md#AboutContributor)

This example deletes the empty rows from a selected range.




```vb
Sub Delete_Empty_Rows()
    'The range from which to delete the rows.
    Dim rnSelection As Range
    
    'Row and count variables used in the deletion process.
    Dim lnLastRow As Long
    Dim lnRowCount As Long
    Dim lnDeletedRows As Long
    
    'Initialize the number of deleted rows.
    lnDeletedRows = 0
    
    'Confirm that a range is selected, and that the range is contiguous.
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            
            'Initialize the range to what the user has selected, and initialize the count for the upcoming FOR loop.
            Set rnSelection = Application.Selection
            lnLastRow = rnSelection.Rows.Count
        
            'Start at the bottom row and work up: if the row is empty then
            'delete the row and increment the deleted row count.
            For lnRowCount = lnLastRow To 1 Step -1
                If Application.CountA(rnSelection.Rows(lnRowCount)) = 0 Then
                    rnSelection.Rows(lnRowCount).Delete
                    lnDeletedRows = lnDeletedRows + 1
                End If
            Next lnRowCount
        
            rnSelection.Resize(lnLastRow - lnDeletedRows).Select
         Else
            MsgBox "Please select only one area.", vbInformation
         End If
    Else
        MsgBox "Please select a range.", vbInformation
    End If
    
    'Turn screen updating back on.
    Application.ScreenUpdating = True

End Sub
```


## About the Contributors
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Range Object](range-object-excel.md)

