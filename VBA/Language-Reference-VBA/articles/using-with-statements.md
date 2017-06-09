---
title: Using With Statements
keywords: vbcn6.chm1076687
f1_keywords:
- vbcn6.chm1076687
ms.prod: office
ms.assetid: ae7f6296-f151-1a1d-a273-a4b80b18b367
ms.date: 06/08/2017
---


# Using With Statements

The  **With** statement lets you specify an[object](vbe-glossary.md) or[user-defined type](vbe-glossary.md) once for an entire series of[statements](vbe-glossary.md).  **With** statements make your procedures run faster and help you avoid repetitive typing.

The following example fills a range of cells with the number 30, applies bold formatting, and sets the interior color of the cells to yellow.



```vb
Sub FormatRange() 
 With Worksheets("Sheet1").Range("A1:C10") 
 .Value = 30 
 .Font.Bold = True 
 .Interior.Color = RGB(255, 255, 0) 
 End With 
End Sub
```

You can nest  **With** statements for greater efficiency. The following example inserts a formula into cell A1, and then formats the font.



```vb
Sub MyInput() 
 With Workbooks("Book1").Worksheets("Sheet1").Cells(1, 1) 
 .Formula = "=SQRT(50)" 
 With .Font 
 .Name = "Arial" 
 .Bold = True 
 .Size = 8 
 End With 
 End With 
End Sub
```


