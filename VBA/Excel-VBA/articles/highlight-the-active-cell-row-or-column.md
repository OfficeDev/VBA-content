---
title: Highlight the Active Cell, Row, or Column
ms.prod: excel
ms.assetid: 51a30ffb-77f2-4bd7-8eb6-b6781dc55d43
ms.date: 06/08/2017
---


# Highlight the Active Cell, Row, or Column

The following code examples show ways to highlight the active cell or the rows and columns that contain the active cell. These examples use the  **[SelectionChange](worksheet-selectionchange-event-excel.md)** event of the **[Worksheet](worksheet-object-excel.md)** object.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)

## Highlighting the Active Cell

The following code example clears the color in all the cells on the worksheet by setting the  **[ColorIndex](interior-colorindex-property-excel.md)** property equal to 0, and then highlights the active cell by setting the **ColorIndex** property equal to 8 (Turquoise).


```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    ' Highlight the active cell
    Target.Interior.ColorIndex = 8
    Application.ScreenUpdating = True
End Sub
```


## Highlighting the Entire Row and Column that Contain the Active Cell

The following code example clears the color in all the cells on the worksheet by setting the  **[ColorIndex](interior-colorindex-property-excel.md)** property equal to 0, and then highlights the entire row and column that contain the active cell by using the **[EntireRow](range-entirerow-property-excel.md)** and **[EntireColumn](range-entirecolumn-property-excel.md)** properties.


```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    With Target
        ' Highlight the entire row and column that contain the active cell
        .EntireRow.Interior.ColorIndex = 8
        .EntireColumn.Interior.ColorIndex = 8
    End With
    Application.ScreenUpdating = True
End Sub
```


## Highlighting the Row and Column that Contain the Active Cell, Within the Current Region

The following code example clears the color in all the cells on the worksheet by setting the  **[ColorIndex](interior-colorindex-property-excel.md)** property equal to 0, and then highlights the row and column that contain the active cell, within the current region by using the **[CurrentRegion](range-currentregion-property-excel.md)** property of the **[Range](range-object-excel.md)** object.


```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    If IsEmpty(Target) Or Target.Cells.Count > 1 Then Exit Sub
    Application.ScreenUpdating = False
    With ActiveCell
        ' Highlight the row and column that contain the active cell, within the current region
        Range(Cells(.Row, .CurrentRegion.Column), Cells(.Row, .CurrentRegion.Columns.Count + .CurrentRegion.Column - 1)).Interior.ColorIndex = 8
        Range(Cells(.CurrentRegion.Row, .Column), Cells(.CurrentRegion.Rows.Count + .CurrentRegion.Row - 1, .Column)).Interior.ColorIndex = 8
    End With
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


