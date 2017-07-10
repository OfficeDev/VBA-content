---
title: Delete Duplicate Entries in a Range
ms.prod: excel
ms.assetid: 22ca07fd-1f69-409a-85e1-247740d87e8e
ms.date: 06/08/2017
---


# Delete Duplicate Entries in a Range

The following example shows how to take a range of data in column A and delete duplicate entries. This example uses the  **[AdvancedFilter](range-advancedfilter-method-excel.md)** method of the **[Range](range-object-excel.md)** object with theUnique parameter equal to **True** to get the unique list of data. TheAction parameter equals **xlFilterInPlace**, specifying that the data is filtered in place. If you want to retain your original data, set the Action parameter equal to **xlFilterCopy** and specify the location where you want the filtered data copied in theCopyToRange parameter. Once the unique values are filtered, this example uses the **[SpecialCells](range-specialcells-method-excel.md)** method of the **Range** object to find any remaining blank rows and deletes them.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)



```vb
Sub DeleteDuplicates()
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = False
        Dim LastColumn As Integer
        LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A1:A" &; Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterInPlace, Unique:=True
            .SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


