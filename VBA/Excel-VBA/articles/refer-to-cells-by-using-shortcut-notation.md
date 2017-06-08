---
title: Refer to Cells by Using Shortcut Notation
keywords: vbaxl10.chm5204430
f1_keywords:
- vbaxl10.chm5204430
ms.prod: excel
ms.assetid: 32426c8d-a2f6-dae5-7507-ff19582fa170
ms.date: 06/08/2017
---


# Refer to Cells by Using Shortcut Notation

You can use either the A1 reference style or a named range within brackets as a shortcut for the  **Range** property. You do not have to type the word "Range" or use quotation marks, as shown in the following examples.


```vb
Sub ClearRange() 
 Worksheets("Sheet1").[A1:B5].ClearContents 
End Sub 
 
Sub SetValue() 
 [MyRange].Value = 30 
End Sub
```


