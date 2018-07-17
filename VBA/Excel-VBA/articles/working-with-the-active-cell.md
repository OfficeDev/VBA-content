---
title: Working with the Active Cell
keywords: vbaxl10.chm5206012
f1_keywords:
- vbaxl10.chm5206012
ms.prod: excel
ms.assetid: 85624b78-b740-6d9b-12cb-b80332c1bf1d
ms.date: 06/08/2017
---


# Working with the Active Cell

The  **[ActiveCell](application-activecell-property-excel.md)** property returns a  **[Range](https://msdn.microsoft.com/en-us/library/office/ff838238.aspx)** object that represents the cell that is active. You can apply any of the properties or methods of a **Range** object to the active cell, as in the following example. While one or more worksheet cells may be selected, only one of the cells in the selection can be the **ActiveCell**.


```vb
Sub SetValue() 
 Worksheets("Sheet1").Activate 
 ActiveCell.Value = 35 
End Sub
```


 **Note**  You can work with the active cell only when the worksheet that it is on is the active sheet.


## Moving the Active Cell

You can use the  **[Range .Activate](https://msdn.microsoft.com/en-us/library/office/aa221681(v=office.11).aspx)** method to designate which cell is the active cell. For example, the following procedure makes B5 the active cell and then formats it as bold.


```vb
Sub SetActive_MakeBold() 
 Worksheets("Sheet1").Activate 
 Worksheets("Sheet1").Range("B5").Activate 
 ActiveCell.Font.Bold = True 
End Sub
```


 **Note**  To select a range of cells, use the  **Select** method. To make a single cell the active cell, use the **Activate** method.

You can use the  **Offset** property to move the active cell. The following procedure inserts text into the active cell in the selected range and then moves the active cell one cell to the right without changing the selection.




```vb
Sub MoveActive() 
 Worksheets("Sheet1").Activate 
 Range("A1:D10").Select 
 ActiveCell.Value = "Monthly Totals" 
 ActiveCell.Offset(0, 1).Activate 
End Sub
```


## Selecting the Cells Surrounding the Active Cell

The  **[CurrentRegion](range-currentregion-property-excel.md)** property returns a range or 'island' of cells bounded by blank rows and columns. In the following example, the selection is expanded to include the cells that contain data immediately adjoining the active cell. This range is then formatted with the Currency style.


```vb
Sub Region() 
 Worksheets("Sheet1").Activate 
 ActiveCell.CurrentRegion.Select 
 Selection.Style = "Currency" 
End Sub
```


