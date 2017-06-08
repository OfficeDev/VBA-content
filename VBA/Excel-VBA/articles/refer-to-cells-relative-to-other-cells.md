---
title: Refer to Cells Relative to Other Cells
keywords: vbaxl10.chm5204424
f1_keywords:
- vbaxl10.chm5204424
ms.prod: excel
ms.assetid: fbdcddea-917c-1813-57a5-21df1c8102de
ms.date: 06/08/2017
---


# Refer to Cells Relative to Other Cells

A common way to work with a cell relative to another cell is to use the  **Offset** property. In the following example, the contents of the cell that is one row down and three columns over from the active cell on the active worksheet are formatted as double-underlined.


```vb
Sub Underline() 
 ActiveCell.Offset(1, 3).Font.Underline = xlDouble 
End Sub
```


 **Note**  You can record macros that use the  **Offset** property to specify relative references instead of absolute references. To do that, on the **Developer** tab, click **Use Relative References**, and then click  **Record Macro**.

To loop through a range of cells, use a variable with the  **Cells** property in a loop. The following example fills the first 20 cells in the third column with values between 5 and 100, incremented by 5. The variable `counter` is used as the row index for the **Cells** property.



```vb
Sub CycleThrough() 
 Dim counter As Integer 
 For counter = 1 To 20 
 Worksheets("Sheet1").Cells(counter, 3).Value = counter * 5 
 Next counter 
End Sub
```


