---
title: Refer to Rows and Columns
keywords: vbaxl10.chm5204439
f1_keywords:
- vbaxl10.chm5204439
ms.prod: excel
ms.assetid: a03acade-9e40-6a26-6a48-2d7a76d0f722
ms.date: 06/08/2017
---


# Refer to Rows and Columns

Use the  **Rows** property or the **Columns** property to work with entire rows or columns. These properties return a **Range** object that represents a range of cells. In the following example, `Rows(1)` returns row one on Sheet1. The **Bold** property of the **Font** object for the range is then set to **True**.


```vb
Sub RowBold() 
    Worksheets("Sheet1").Rows(1).Font.Bold = True 
End Sub
```


The following table illustrates some row and column references using the  **Rows** and **Columns** properties.



|**Reference**|**Meaning**|
|:-----|:-----|
| `Rows(1)`|Row one|
| `Rows`|All the rows on the worksheet|
| `Columns(1)`|Column one|
| `Columns("A")`|Column one|
| `Columns`|All the columns on the worksheet|
To work with several rows or columns at the same time, create an object variable and use the  **Union** method, combining multiple calls to the **Rows** or **Columns** property. The following example changes the format of rows one, three, and five on worksheet one in the active workbook to bold.



```vb
Sub SeveralRows() 
    Worksheets("Sheet1").Activate 
    Dim myUnion As Range 
    Set myUnion = Union(Rows(1), Rows(3), Rows(5)) 
    myUnion.Font.Bold = True 
End Sub
```

 **Sample code provided by:** Dennis Wallentin, [VSTO &; .NET &; Excel](http://xldennis.wordpress.com/)
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

This example deletes the empty columns from a selected range.



```vb
Sub Delete_Empty_Columns()
    'The range from which to delete the columns.
    Dim rnSelection As Range
    
    'Column and count variables used in the deletion process.
    Dim lnLastColumn As Long
    Dim lnColumnCount As Long
    Dim lnDeletedColumns As Long
    
    lnDeletedColumns = 0
    
    'Confirm that a range is selected, and that the range is contiguous.
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            
            'Initialize the range to what the user has selected, and initialize the count for the upcoming FOR loop.
            Set rnSelection = Application.Selection
            lnLastColumn = rnSelection.Columns.Count
        
            'Start at the far-right column and work left: if the column is empty then
            'delete the column and increment the deleted column count.
            For lnColumnCount = lnLastColumn To 1 Step -1
                If Application.CountA(rnSelection.Columns(lnColumnCount)) = 0 Then
                    rnSelection.Columns(lnColumnCount).Delete
                    lnDeletedColumns = lnDeletedColumns + 1
                End If
            Next lnColumnCount
    
            rnSelection.Resize(lnLastColumn - lnDeletedColumns).Select
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


## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &; .NET &; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


