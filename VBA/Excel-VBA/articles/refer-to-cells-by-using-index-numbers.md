---
title: Refer to Cells by Using Index Numbers
keywords: vbaxl10.chm5204428
f1_keywords:
- vbaxl10.chm5204428
ms.prod: excel
ms.assetid: 5671563b-9a20-3124-58d9-cfa02fac5312
ms.date: 06/08/2017
---


# Refer to Cells by Using Index Numbers

You can use the  **Cells** property to refer to a single cell by using row and column index numbers. This property returns a  **Range** object that represents a single cell. In the following example, `Cells(6,1)` returns cell A6 on Sheet1. The **Value** property is then set to 10.


```vb
Sub EnterValue() 
 Worksheets("Sheet1").Cells(6, 1).Value = 10 
End Sub
```


The  **Cells** property works well for looping through a range of cells, because you can substitute variables for the index numbers, as shown in the following example.




```vb
Sub CycleThrough() 
 Dim Counter As Integer 
 For Counter = 1 To 20 
 Worksheets("Sheet1").Cells(Counter, 3).Value = Counter 
 Next Counter 
End Sub
```


 **Note**  If you want to change the properties of (or apply a method to) a range of cells all at once, use the  **Range** property. For more information, see [Refer to Cells and Ranges by Using A1 Notation](refer-to-cells-and-ranges-by-using-a1-notation.md).


