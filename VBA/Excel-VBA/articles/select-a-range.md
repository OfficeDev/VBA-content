---
title: Select a Range
ms.prod: excel
ms.assetid: 4ec2e533-74b3-448d-90aa-1e2a624490b8
ms.date: 06/08/2017
---


# Select a Range

These examples show how to select the used range, which includes formatted cells that do not contain data, and how to select a data range, which includes cells that contains actual data.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)

## Selecting the Used Range

This example shows how to select the used range on the current sheet, which includes formatted cells that do not contain data, by using the  **[UsedRange](worksheet-usedrange-property-excel.md)** property of the **[Worksheet](worksheet-object-excel.md)** object and the **[Select](range-select-method-excel.md)** method of the **[Range](range-object-excel.md)** object. Then it displays the address of the range to the user.


```vb
Sub SelectUsedRange()
    ActiveSheet.UsedRange.Select
    MsgBox "The used range address is " &; ActiveSheet.UsedRange.Address(0, 0) &; ".", 64, "Used range address:"
End Sub
```


## Selecting a Data Range Starting at Cell A1

This example shows how to select a data range on the current sheet, starting at cell A1, and display the address of the range to the user. The data range does not include cells that are formatted that do not contain data. To get the data range, this example finds the last row and the last column that contain actual data by using the  **[Find](range-find-method-excel.md)** method of the **[Range](range-object-excel.md)** object.


```vb
Sub SelectDataRange()
    Dim LastRow As Long, LastColumn As Long
    LastRow = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Range("A1").Resize(LastRow, LastColumn).Select
    MsgBox "The data range address is " &; Selection.Address(0, 0) &; ".", 64, "Data-containing range address:"
End Sub
```


## Selecting a Data Range of Unknown Starting Location

This example shows how to select a data range on the current sheet when you do not know the starting location, and display the address of the range to the user. The data range does not include cells that are formatted that do not contain data. To get the data range, this example finds the first and last row and column that contain actual data by using the  **[Find](range-find-method-excel.md)** method of the **[Range](range-object-excel.md)** object.


```vb
Sub UnknownRange()
    If WorksheetFunction.CountA(Cells) = 0 Then
        MsgBox "There is no range to be selected.", , "No cells contain any values."
        Exit Sub
    Else
        Dim FirstRow&;, FirstCol&;, LastRow&;, LastCol&;
        Dim myUsedRange As Range
        FirstRow = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
        
        On Error Resume Next
        FirstCol = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
        If Err.Number <> 0 Then
            Err.Clear
            MsgBox _
            "There are horizontally merged cells on the sheet" &; vbCrLf &; _
            "that should be removed in order to locate the range.", 64, "Please unmerge all cells."
            Exit Sub
        End If
        
        LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
        LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
        Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
        myUsedRange.Select
        MsgBox "The data range on this worksheet is " &; myUsedRange.Address(0, 0) &; ".", 64, "Range address:"
    End If
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


