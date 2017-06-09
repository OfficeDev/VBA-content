---
title: Refer to Cells by Using a Range Object
keywords: vbaxl10.chm5204426
f1_keywords:
- vbaxl10.chm5204426
ms.prod: excel
ms.assetid: 89c2d61d-823a-9376-d827-2ec5ae200d80
ms.date: 06/08/2017
---


# Refer to Cells by Using a Range Object

If you set an object variable to a  **Range** object, you can easily manipulate the range by using the variable name.

The following procedure creates the object variable  `myRange` and then assigns the variable to range A1:D5 on Sheet1 in the active workbook. Subsequent statements modify properties of the range by substituting the variable name for the **Range** object.



```vb
Sub Random() 
 Dim myRange As Range 
 Set myRange = Worksheets("Sheet1").Range("A1:D5") 
 myRange.Formula = "=RAND()" 
 myRange.Font.Bold = True 
End Sub
```


