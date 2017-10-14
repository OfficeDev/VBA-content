---
title: Looping Through a Range of Cells
keywords: vbaxl10.chm5202755
f1_keywords:
- vbaxl10.chm5202755
ms.prod: excel
ms.assetid: ee134d2e-851d-eaaa-009a-90fff7db7517
ms.date: 06/08/2017
---


# Looping Through a Range of Cells

When using Visual Basic, you often need to run the same block of statements on each cell in a range of cells. To do this, you combine a looping statement and one or more methods to identify each cell, one at a time, and run the operation.

One way to loop through a range is to use the  **For...Next** loop with the **Cells** property. Using the **Cells** property, you can substitute the loop counter (or other variables or expressions) for the cell index numbers. In the following example, the variable `counter` is substituted for the row index. The procedure loops through the range C1:C20, setting to 0 (zero) any number whose absolute value is less than 0.01.



```vb
Sub RoundToZero1() 
 For Counter = 1 To 20 
 Set curCell = Worksheets("Sheet1").Cells(Counter, 3) 
 If Abs(curCell.Value) < 0.01 Then curCell.Value = 0 
 Next Counter 
End Sub
```

Another easy way to loop through a range is to use a  **For Each...Next** loop with the collection of cells specified in the **Range** property. Visual Basic automatically sets an object variable for the next cell each time the loop runs. The following procedure loops through the range A1:D10, setting to 0 (zero) any number whose absolute value is less than 0.01.



```vb
Sub RoundToZero2() 
 For Each c In Worksheets("Sheet1").Range("A1:D10").Cells 
 If Abs(c.Value) < 0.01 Then c.Value = 0 
 Next 
End Sub
```

If you do not know the boundaries of the range you want to loop through, you can use the  **CurrentRegion** property to return the range that surrounds the active cell. For example, the following procedure, when run from a worksheet, loops through the range that surrounds the active cell, setting to 0 (zero) any number whose absolute value is less than 0.01.



```vb
Sub RoundToZero3() 
 For Each c In ActiveCell.CurrentRegion.Cells 
 If Abs(c.Value) < 0.01 Then c.Value = 0 
 Next 
End Sub
```


