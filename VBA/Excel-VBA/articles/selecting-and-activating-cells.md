---
title: Selecting and Activating Cells
keywords: vbaxl10.chm5204866
f1_keywords:
- vbaxl10.chm5204866
ms.prod: excel
ms.assetid: bdfead4b-0909-e67d-e478-7fb33aceec79
ms.date: 06/08/2017
---


# Selecting and Activating Cells

In Microsoft Excel, you usually select a cell or cells and then perform an action, such as formatting the cells or entering values in them. In Visual Basic, it is usually not necessary to select cells before modifying them.

For example, to enter a formula in cell D6 using Visual Basic, you do not need to select the range D6. Just return the  **Range** object for that cell, and then set the **Formula** property to the formula you want, as shown in the following example.



```VB.net
Sub EnterFormula() 
    Worksheets("Sheet1").Range("D6").Formula = "=SUM(D2:D5)" 
End Sub
```

For more information and examples of using other methods to control cells without selecting them, see  [How to: Reference Cells and Ranges](reference-cells-and-ranges.md).

## Using the Select Method and the Selection Property

The  **Select** method activates sheets and objects on sheets; the **Selection** property returns an object that represents the current selection on the active sheet in the active workbook. Before you can use the **Selection** property successfully, you must activate a workbook, activate or select a sheet, and then select a range (or other object) using the **Select** method.

The macro recorder will often create a macro that uses the  **Select** method and the **Selection** property. The following **Sub** procedure was created using the macro recorder, and it shows how **Select** and **Selection** work together.




```VB.net
Sub Macro1() 
    Sheets("Sheet1").Select 
    Range("A1").Select 
    ActiveCell.FormulaR1C1 = "Name" 
    Range("B1").Select 
    ActiveCell.FormulaR1C1 = "Address" 
    Range("A1:B1").Select 
    Selection.Font.Bold = True 
End Sub
```

The following example performs the same task without activating or selecting the worksheet or cells.




```VB.net
Sub Labels() 
    With Worksheets("Sheet1") 
        .Range("A1") = "Name" 
        .Range("B1") = "Address" 
        .Range("A1:B1").Font.Bold = True 
    End With 
End Sub
```


## Selecting Cells on the Active Worksheet

If you use the  **Select** method to select cells, be aware that **Select** works only on the active worksheet. If you run your **Sub** procedure from the module, the **Select** method will fail unless your procedure activates the worksheet before using the **Select** method on a range of cells. For example, the following procedure copies a row from Sheet1 to Sheet2 in the active workbook.


```VB.net
Sub CopyRow() 
    Worksheets("Sheet1").Rows(1).Copy 
    Worksheets("Sheet2").Select 
    Worksheets("Sheet2").Rows(1).Select 
    Worksheets("Sheet2").Paste 
End Sub
```


## Activating a Cell Within a Selection

You can use the  **Activate** method to activate a cell within a selection. There can be only one active cell, even when a range of cells is selected. The following procedure selects a range and then activates a cell within the range without changing the selection.


```VB.net
Sub MakeActive() 
    Worksheets("Sheet1").Activate 
    Range("A1:D4").Select 
    Range("B2").Activate 
End Sub
```


